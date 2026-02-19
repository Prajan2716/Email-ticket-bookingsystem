<?php
/**
 * TicketSyncer - Core Gmail → Sheets sync logic
 * PHP equivalent of main.py
 */

declare(strict_types=1);

namespace GmailTicket;

use Google\Service\Gmail;

class TicketSyncer
{
    // ──────────────────────────────────────────────────────
    //  Config
    // ──────────────────────────────────────────────────────

    private const SPREADSHEET_ID            = '1F1STl7ubwviajSvu1LGfw6FGiDUQJbcZ-fI586pqwiE';
    private const WORKSHEET                 = 'Email log';
    private const THREAD_STATE_FILE         = __DIR__ . '/../thread_state.txt';
    private const SHEET_BACKUP_INTERVAL     = 50;   // Back up to sheet every N syncs
    private const TICKET_MAP_REFRESH_INTERVAL = 20; // Refresh ticket map every N syncs

    // Auto-close settings
    private const AUTO_CLOSE_ENABLED = true;
    private const AUTO_CLOSE_HOURS   = 6;
    private const AUTO_CLOSE_ACTION  = 'close'; // 'close' or 'delete'

    // Hard-coded admin emails (merged with sheet at runtime)
    private const BASE_ADMIN_EMAILS = [
        'support-ticketana@he5.in',
        // Add more admin emails here
    ];

    // ──────────────────────────────────────────────────────
    //  Runtime state (equivalent to Python globals)
    // ──────────────────────────────────────────────────────

    private int   $syncCounter            = 0;
    private ?array $cachedThreadMap       = null;  // thread_id => row_number
    private int   $lastTicketMapRefresh   = 0;
    private bool  $sheetsInitialized      = false;

    /** @var array<string, true>  email => true */
    private array $knownSenders       = [];
    private bool  $knownSendersLoaded = false;

    /** @var string[]  All admin emails (base + sheet) */
    private array $adminEmails = self::BASE_ADMIN_EMAILS;

    private GmailHandler $gmailHandler;
    private SheetsHandler $sheetsHandler;

    // ──────────────────────────────────────────────────────
    //  Construction
    // ──────────────────────────────────────────────────────

    public function __construct()
    {
        $this->gmailHandler  = new GmailHandler();
        $gmailClient         = $this->gmailHandler->getClient();
        $this->sheetsHandler = new SheetsHandler($gmailClient);
    }

    // ──────────────────────────────────────────────────────
    //  File-based thread state
    // ──────────────────────────────────────────────────────

    /**
     * @return array<string, int>  thread_id => timestamp
     */
    private function loadThreadStateFromFile(): array
    {
        $state = [];
        if (!file_exists(self::THREAD_STATE_FILE)) {
            return $state;
        }

        foreach (file(self::THREAD_STATE_FILE, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES) as $line) {
            if (str_contains($line, '|')) {
                [$tid, $ts] = explode('|', $line, 2);
                $state[trim($tid)] = (int) trim($ts);
            }
        }

        return $state;
    }

    /**
     * @param array<string, int> $state
     */
    private function saveThreadStateToFile(array $state): void
    {
        $lines = [];
        foreach ($state as $tid => $ts) {
            $lines[] = "{$tid}|{$ts}";
        }
        file_put_contents(self::THREAD_STATE_FILE, implode("\n", $lines) . "\n");
    }

    // ──────────────────────────────────────────────────────
    //  Admin / sender helpers
    // ──────────────────────────────────────────────────────

    private function isAdminEmail(string $email): bool
    {
        return in_array(strtolower($email), $this->adminEmails, true);
    }

    private function isNewSenderCached(string $fromEmail): bool
    {
        return !isset($this->knownSenders[strtolower($fromEmail)]);
    }

    // ──────────────────────────────────────────────────────
    //  Auto-close stale tickets
    // ──────────────────────────────────────────────────────

    private function checkAndCloseStaleTickets(Gmail $gmail): void
    {
        if (!self::AUTO_CLOSE_ENABLED) {
            return;
        }

        echo "\n🔍 Checking for stale tickets (>" . self::AUTO_CLOSE_HOURS . "h no customer response)...\n";

        $allRows    = $this->sheetsHandler->getAllRows(self::SPREADSHEET_ID, self::WORKSHEET);
        $currentTs  = time();
        $closedCount  = 0;
        $deletedCount = 0;

        // If "delete" action, we need the numeric sheet ID
        $sheetId = null;
        if (self::AUTO_CLOSE_ACTION === 'delete') {
            $sheetId = $this->sheetsHandler->getSheetId(self::SPREADSHEET_ID, self::WORKSHEET);
        }

        // Track deleted rows so we can adjust row numbers
        $deletedRows = [];

        foreach ($allRows as $idx => $row) {
            if ($idx === 0 || count($row) < 6) {
                continue; // Skip header and short rows
            }

            $rowNumber   = $idx + 1; // 1-indexed
            $ticketId    = $row[0] ?? '';
            $threadId    = $row[1] ?? '';
            $timestampStr = $row[2] ?? '';
            $status      = $row[5] ?? '';

            if ($status !== 'Awaiting customer reply') {
                continue;
            }

            // Parse the ticket timestamp
            $ticketTimestamp = strtotime($timestampStr);
            if ($ticketTimestamp === false) {
                continue;
            }

            $hoursPassed = ($currentTs - $ticketTimestamp) / 3600.0;

            if ($hoursPassed < self::AUTO_CLOSE_HOURS) {
                continue;
            }

            if (self::AUTO_CLOSE_ACTION === 'delete') {
                // Adjust row number for previously deleted rows
                $adjustedRow = $rowNumber - count(array_filter($deletedRows, fn($r) => $r < $rowNumber));

                $this->sheetsHandler->deleteRow(self::SPREADSHEET_ID, $sheetId, $adjustedRow);
                $deletedRows[] = $rowNumber;

                echo sprintf("   🗑️  Deleted ticket %s (no response for %.1fh)\n", $ticketId, $hoursPassed);

                // Trash the Gmail thread
                try {
                    $this->gmailHandler->trashThread($gmail, $threadId);
                    echo "   📧 Trashed email thread {$threadId}\n";
                } catch (\Exception $e) {
                    echo "   ⚠️  Could not trash thread: {$e->getMessage()}\n";
                }

                $deletedCount++;
            } else {
                // Mark as closed
                $row[5] = 'Closed - No customer response';
                $this->sheetsHandler->updateRow(self::SPREADSHEET_ID, self::WORKSHEET, $rowNumber, $row);
                echo sprintf("   ✅ Closed ticket %s (no response for %.1fh)\n", $ticketId, $hoursPassed);
                $closedCount++;
            }
        }

        if ($closedCount > 0 || $deletedCount > 0) {
            echo "📊 Auto-close summary: {$closedCount} closed, {$deletedCount} deleted\n";
        }
    }

    // ──────────────────────────────────────────────────────
    //  Main sync
    // ──────────────────────────────────────────────────────

    public function syncMailToSheet(): void
    {
        echo "\n" . str_repeat('=', 50) . "\n";
        echo "Starting sync...\n";
        echo str_repeat('=', 50) . "\n";

        // ── Gmail setup ──────────────────────────────────────
        $gmail   = $this->gmailHandler->getService();
        $myEmail = $this->gmailHandler->getMyEmail($gmail);

        // Ensure authenticated email is in admin list
        $myEmailLower = strtolower($myEmail);
        if (!in_array($myEmailLower, $this->adminEmails, true)) {
            $this->adminEmails[] = $myEmailLower;
        }

        echo "📧 Authenticated as: {$myEmail}\n";
        echo "👥 Admin emails: " . implode(', ', $this->adminEmails) . "\n";

        // ── Sheets setup ─────────────────────────────────────
        echo "📊 Connected to spreadsheet\n";

        // One-time initialisation of state worksheets
        if (!$this->sheetsInitialized) {
            $this->sheetsHandler->initializeStateSheets(self::SPREADSHEET_ID);

            // Load admin emails from sheet (merging with hardcoded list)
            //$sheetAdmins       = $this->sheetsHandler->loadAdminEmailsFromSheet(self::SPREADSHEET_ID);
            //$this->adminEmails = array_unique(array_merge($this->adminEmails, $sheetAdmins));
            //echo "📧 Loaded " . count($sheetAdmins) . " admin emails from sheet\n";

            $this->sheetsInitialized = true;
        }

        // ── Gmail labels ─────────────────────────────────────
        $adminLabel   = $this->gmailHandler->getOrCreateLabel($gmail, 'Awaiting_Admin_Reply');
        $custLabel    = $this->gmailHandler->getOrCreateLabel($gmail, 'Awaiting_Customer_Reply');
        $noreplyLabel = $this->gmailHandler->getOrCreateLabel($gmail, 'No_Reply_Mail');
        echo "🏷️  Labels configured\n";

        // ── Ticket map (cached) ───────────────────────────────
        $this->syncCounter++;

        if (
            $this->cachedThreadMap === null ||
            ($this->syncCounter - $this->lastTicketMapRefresh) >= self::TICKET_MAP_REFRESH_INTERVAL
        ) {
            $this->cachedThreadMap      = $this->sheetsHandler->getAllTickets(self::SPREADSHEET_ID, self::WORKSHEET);
            $this->lastTicketMapRefresh = $this->syncCounter;
            echo '📋 Refreshed ticket map: ' . count($this->cachedThreadMap) . " existing tickets\n";
        } else {
            echo '📋 Using cached ticket map: ' . count($this->cachedThreadMap) . " existing tickets\n";
        }

        // ── Known senders (loaded once) ───────────────────────
        if (!$this->knownSendersLoaded) {
            $this->knownSenders       = $this->sheetsHandler->loadKnownSenders(self::SPREADSHEET_ID, self::WORKSHEET);
            $this->knownSendersLoaded = true;
            echo '📧 Loaded ' . count($this->knownSenders) . " known senders from sheet\n";
        }

        // ── Thread state (file-based, fast) ───────────────────
        $threadState = $this->loadThreadStateFromFile();

        // First-run: load last-sync timestamp from sheet; subsequent runs use current time - 10s
        if ($this->syncCounter === 1) {
            $lastSync = $this->sheetsHandler->getLastSyncTimestamp(self::SPREADSHEET_ID);
            echo "📊 Loaded initial sync timestamp from sheet\n";
        } else {
            $lastSync = time() - 10; // Look back 10 seconds to be safe
        }

        echo '📊 Loaded state: ' . count($threadState) . " threads tracked (sync #{$this->syncCounter})\n";

        // ── Build query and fetch threads ─────────────────────
        $query   = $lastSync ? "after:{$lastSync}" : 'newer_than:7d';
        $threads = $this->gmailHandler->fetchAllThreads($gmail, $query);

        // Deduplicate (Gmail sometimes returns the same thread twice)
        $seen          = [];
        $uniqueThreads = [];
        foreach ($threads as $t) {
            if (!isset($seen[$t['id']])) {
                $seen[$t['id']] = true;
                $uniqueThreads[] = $t;
            }
        }

        $removed = count($threads) - count($uniqueThreads);
        if ($removed > 0) {
            echo "⚠️  Removed {$removed} duplicate thread(s)\n";
        }

        $threads = $uniqueThreads;
        echo '📬 Found ' . count($threads) . " threads to process\n\n";

        // ── Process each thread ───────────────────────────────
        foreach ($threads as $threadInfo) {
            $tid = $threadInfo['id'];

            echo "\n" . str_repeat('=', 60) . "\n";
            echo "🔍 DEBUG: Examining thread {$tid}\n";
            echo '   In cachedThreadMap: ' . (isset($this->cachedThreadMap[$tid]) ? 'yes' : 'no') . "\n";
            echo '   In threadState: '     . (isset($threadState[$tid]) ? 'yes' : 'no') . "\n";
            if (isset($threadState[$tid])) {
                echo "   threadState timestamp: {$threadState[$tid]}\n";
            }
            echo str_repeat('=', 60) . "\n";

            // Get full thread
            $thread              = $this->gmailHandler->getThreadDetails($gmail, $tid);
            [$msg, $headers]     = $this->gmailHandler->getLastMessage($thread);

            if ($msg === null) {
                echo "⏭️  Skipping thread {$tid} - no messages\n";
                continue;
            }

            $ts = (int) ($msg->getInternalDate() / 1000);

            // Skip if already processed
            if ($ts <= ($threadState[$tid] ?? 0)) {
                echo "⏭️  Skipping thread {$tid} - already processed\n";
                continue;
            }

            $fromEmail = $this->gmailHandler->extractEmail($headers['From'] ?? '');
            $subject   = $headers['Subject'] ?? 'No Subject';

            echo "\n📨 Processing thread {$tid}\n";
            echo "   From: {$fromEmail}\n";
            echo "   Subject: {$subject}\n";

            $isNoreply   = $this->gmailHandler->isNoreplyEmail($fromEmail);
            $isNewTicket = !isset($this->cachedThreadMap[$tid]);

            if ($isNoreply) {
                echo "   🚫 NO-REPLY EMAIL DETECTED: {$fromEmail}\n";
            }
            echo '   🎯 DEBUG: isNewTicket = ' . ($isNewTicket ? 'true' : 'false') . "\n";
            echo '   🎯 DEBUG: cachedThreadMap size = ' . count($this->cachedThreadMap) . "\n";

            // Skip if already fully processed in this sync cycle
            if (isset($threadState[$tid]) && $threadState[$tid] >= $ts) {
                echo "   ⏭️  Skipping thread {$tid} - already processed in this sync\n";
                continue;
            }

            // ── No-reply on NEW ticket ────────────────────────
            if ($isNoreply && $isNewTicket) {
                echo "   🚫 Processing no-reply email as closed ticket\n";

                // Final safety check (race-condition guard)
                $finalCheckMap = $this->sheetsHandler->getAllTickets(self::SPREADSHEET_ID, self::WORKSHEET);
                if (isset($finalCheckMap[$tid])) {
                    echo "   ⚠️  WARNING: Thread {$tid} was just created by another process!\n";
                    echo "   ⏭️  Skipping to avoid duplicate\n";
                    $this->cachedThreadMap = $finalCheckMap;
                    $threadState[$tid]     = $ts;
                    continue;
                }

                $ticketId = $this->sheetsHandler->getNextTicketId(self::SPREADSHEET_ID);
                echo "   🎫 New no-reply ticket: {$ticketId}\n";

                $rowData = $this->sheetsHandler->createTicketRow(
                    $ticketId, $tid, $fromEmail, $subject, 'No-reply - Closed', false
                );
                $this->sheetsHandler->addNewTicket(self::SPREADSHEET_ID, self::WORKSHEET, $rowData);
                echo "   ✅ Created no-reply ticket {$ticketId}\n";

                $this->gmailHandler->updateThreadLabels($gmail, $tid, [$noreplyLabel], [$adminLabel, $custLabel]);
                echo "   🏷️  Added 'No_Reply_Mail' label to thread\n";

                $threadState[$tid]     = $ts;
                $this->cachedThreadMap = $this->sheetsHandler->getAllTickets(self::SPREADSHEET_ID, self::WORKSHEET);
                echo "   🛑 Thread stopped - no further updates will be processed\n";
                continue;
            }

            // ── No-reply on EXISTING ticket ───────────────────
            if ($isNoreply && !$isNewTicket) {
                echo "   ⏭️  Skipping - no-reply email on existing ticket\n";
                $threadState[$tid] = $ts;
                continue;
            }

            // ── Skip NEW threads initiated by admin ───────────
            if ($isNewTicket && $this->isAdminEmail($fromEmail)) {
                echo "   ⏭️  Skipping - admin-initiated thread\n";
                $threadState[$tid] = $ts;
                continue;
            }

            // ── Resolve or create ticket ID ───────────────────
            if (!$isNewTicket) {
                $rowNum     = $this->cachedThreadMap[$tid];
                $ticketData = $this->sheetsHandler->getRow(self::SPREADSHEET_ID, self::WORKSHEET, $rowNum);
                $ticketId   = $ticketData[0] ?? '';
            } else {
                // Race-condition safety check
                $finalCheckMap = $this->sheetsHandler->getAllTickets(self::SPREADSHEET_ID, self::WORKSHEET);
                if (isset($finalCheckMap[$tid])) {
                    echo "   ⚠️  WARNING: Thread {$tid} was just created by another process!\n";
                    echo "   ⏭️  Skipping to avoid duplicate\n";
                    $this->cachedThreadMap = $finalCheckMap;
                    continue;
                }

                $ticketId = $this->sheetsHandler->getNextTicketId(self::SPREADSHEET_ID);
                echo "   🎫 New ticket: {$ticketId}\n";
                echo "   🆔 DEBUG: Full thread ID = {$tid}\n";
                echo "   🆔 DEBUG: Thread ID length = " . strlen($tid) . "\n";

                // Mark as processed BEFORE creating to prevent duplicate creation
                $threadState[$tid] = $ts;
                echo "   ✅ DEBUG: Marked {$tid} as processed with timestamp {$ts}\n";
            }

            // ── Status & labels ───────────────────────────────
            $status = $this->isAdminEmail($fromEmail) ? 'Awaiting customer reply' : 'Awaiting admin reply';

            $labelsToAdd    = ($status === 'Awaiting admin reply') ? [$adminLabel] : [$custLabel];
            $labelsToRemove = ($status === 'Awaiting admin reply') ? [$custLabel]  : [$adminLabel];

            $this->gmailHandler->updateThreadLabels($gmail, $tid, $labelsToAdd, $labelsToRemove);

            // ── New-sender detection ──────────────────────────
            $newSender = false;
            if ($isNewTicket) {
                $newSender = $this->isNewSenderCached($fromEmail);
                if ($newSender) {
                    echo "   🆕 NEW SENDER: {$fromEmail}\n";
                    $this->knownSenders[strtolower($fromEmail)] = true;
                } else {
                    echo "   👤 Known sender: {$fromEmail}\n";
                }
            }

            // ── Write to sheet ────────────────────────────────
            $rowData = $this->sheetsHandler->createTicketRow(
                $ticketId, $tid, $fromEmail, $subject, $status, $newSender
            );

            if (!$isNewTicket) {
                $this->sheetsHandler->updateExistingTicket(
                    self::SPREADSHEET_ID, self::WORKSHEET, $rowNum, $rowData
                );
                echo "   ✅ Updated ticket {$ticketId}\n";
            } else {
                $this->sheetsHandler->addNewTicket(self::SPREADSHEET_ID, self::WORKSHEET, $rowData);
                echo "   ✅ Created ticket {$ticketId}\n";

                // Refresh cache immediately after creating a new ticket
                $this->cachedThreadMap = $this->sheetsHandler->getAllTickets(self::SPREADSHEET_ID, self::WORKSHEET);
                echo "   🔄 Refreshed cache to include new ticket\n";
            }

            // Update processed timestamp
            $threadState[$tid] = $ts;
        }

        // ── Save state ────────────────────────────────────────
        if (!empty($threads)) {
            $this->saveThreadStateToFile($threadState);
            echo "💾 Saved thread state to file\n";
        }

        // Periodic backup to sheet
        if ($this->syncCounter % self::SHEET_BACKUP_INTERVAL === 0) {
            $this->sheetsHandler->saveThreadStateToSheet(self::SPREADSHEET_ID, $threadState);
            $this->sheetsHandler->saveLastSyncTimestamp(self::SPREADSHEET_ID, time());
            echo "📊 Backed up thread state AND sync timestamp to sheet (sync #{$this->syncCounter})\n";
        }

        // Periodic stale-ticket check
        if ($this->syncCounter % 20 === 0) {
            $this->checkAndCloseStaleTickets($gmail);
        }

        echo "\n" . str_repeat('=', 50) . "\n";
        echo "✅ Sync complete!\n";
        echo str_repeat('=', 50) . "\n\n";
    }
}
