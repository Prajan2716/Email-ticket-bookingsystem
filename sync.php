<?php
/**
 * sync.php — CLI entry point for the background sync
 *
 * Run via cron every minute:
 *   * * * * * /usr/bin/php /path/to/project/sync.php >> /var/log/gmail-ticket.log 2>&1
 *
 * Or run once manually:
 *   php sync.php
 *
 * The script loops internally (just like Flask-APScheduler's 5-second interval)
 * for the duration defined by LOOP_DURATION_SECONDS, then exits so the cron
 * job can restart it cleanly.
 */

declare(strict_types=1);

require_once __DIR__ . '/vendor/autoload.php';

use GmailTicket\TicketSyncer;

// ── Config ────────────────────────────────────────────────────────────────────

const SYNC_INTERVAL_SECONDS  = 5;    // How often to sync (matches original 5-second interval)
const LOOP_DURATION_SECONDS  = 55;   // Run for ~55 s then let cron restart (fits in 1-min cron slot)
const STATUS_FILE            = __DIR__ . '/sync_status.json';

// ── Helpers ───────────────────────────────────────────────────────────────────

function updateSyncStatus(?string $error, int &$syncCount): void
{
    $current = [];
    if (file_exists(STATUS_FILE)) {
        $current = json_decode(file_get_contents(STATUS_FILE), true) ?? [];
    }

    $syncCount = ($current['sync_count'] ?? 0) + 1;

    $current['last_sync']  = date('Y-m-d H:i:s');
    $current['last_error'] = $error;
    $current['sync_count'] = $syncCount;

    file_put_contents(STATUS_FILE, json_encode($current, JSON_PRETTY_PRINT));
}

// ── Lock file (prevent overlapping cron runs) ─────────────────────────────────

$lockFile = sys_get_temp_dir() . '/gmail_ticket_sync.lock';
$lockFp   = fopen($lockFile, 'c');

if (!flock($lockFp, LOCK_EX | LOCK_NB)) {
    echo "⚠️  Another sync process is already running. Exiting.\n";
    exit(0);
}

// ── Main loop ─────────────────────────────────────────────────────────────────

$syncer    = new TicketSyncer();
$startTime = time();
$syncCount = 0;

echo "🚀 Gmail Ticket Sync started at " . date('Y-m-d H:i:s') . "\n";
echo "🔄 Will sync every " . SYNC_INTERVAL_SECONDS . "s for " . LOOP_DURATION_SECONDS . "s, then exit\n\n";

while ((time() - $startTime) < LOOP_DURATION_SECONDS) {
    $cycleStart = microtime(true);

    echo "\n⏰ Scheduled sync triggered at " . date('Y-m-d H:i:s') . "\n";

    try {
        $syncer->syncMailToSheet();
        updateSyncStatus(null, $syncCount);
    } catch (\Throwable $e) {
        $errorMsg = $e->getMessage();
        echo "❌ Sync error: {$errorMsg}\n";
        updateSyncStatus($errorMsg, $syncCount);
    }

    $elapsed = microtime(true) - $cycleStart;
    $sleep   = max(0, SYNC_INTERVAL_SECONDS - (int) $elapsed);

    if ($sleep > 0 && (time() - $startTime + $sleep) < LOOP_DURATION_SECONDS) {
        echo "💤 Sleeping {$sleep}s until next sync...\n";
        sleep($sleep);
    }
}

echo "\n🏁 Sync loop finished after " . (time() - $startTime) . "s. Total syncs this run: {$syncCount}\n";

// Release lock
flock($lockFp, LOCK_UN);
fclose($lockFp);
