<?php
/**
 * SheetsHandler - Manages all Google Sheets operations
 */

declare(strict_types=1);

namespace GmailTicket;

use Google\Client;
use Google\Service\Sheets;
use Google\Service\Sheets\ValueRange;
use Google\Service\Sheets\Spreadsheet;
use Google\Service\Sheets\Sheet;
use Google\Service\Sheets\Request;
use Google\Service\Sheets\BatchUpdateSpreadsheetRequest;
use Google\Service\Sheets\AddSheetRequest;
use Google\Service\Sheets\SheetProperties;

class SheetsHandler
{
    private Sheets $service;

    public function __construct(Client $client)
    {
        $this->service = new Sheets($client);
    }

    // ──────────────────────────────────────────────────────
    //  Spreadsheet / worksheet helpers
    // ──────────────────────────────────────────────────────

    /**
     * Read all values from a worksheet range.
     *
     * @return array<int, array<int, string>>
     */
    public function getValues(string $spreadsheetId, string $range): array
    {
        $response = $this->service->spreadsheets_values->get($spreadsheetId, $range);
        return $response->getValues() ?? [];
    }

    /**
     * Write a 2-D array to a range (USER_ENTERED so formulas are evaluated).
     *
     * @param array<int, array<int, mixed>> $values
     */
    public function setValues(string $spreadsheetId, string $range, array $values): void
    {
        $body = new ValueRange(['values' => $values]);
        $this->service->spreadsheets_values->update(
            $spreadsheetId,
            $range,
            $body,
            ['valueInputOption' => 'USER_ENTERED']
        );
    }

    /**
     * Append a single row (USER_ENTERED).
     *
     * @param array<int, mixed> $row
     */
    public function appendRow(string $spreadsheetId, string $range, array $row): void
    {
        $body = new ValueRange(['values' => [$row]]);
        $this->service->spreadsheets_values->append(
            $spreadsheetId,
            $range,
            $body,
            ['valueInputOption' => 'USER_ENTERED']
        );
    }

    /**
     * Read a single cell value.
     */
    public function getCell(string $spreadsheetId, string $cell): ?string
    {
        $values = $this->getValues($spreadsheetId, $cell);
        return $values[0][0] ?? null;
    }

    /**
     * Write a single cell value.
     */
    public function setCell(string $spreadsheetId, string $cell, mixed $value): void
    {
        $this->setValues($spreadsheetId, $cell, [[$value]]);
    }

    /**
     * Clear all values in a range.
     */
    public function clearRange(string $spreadsheetId, string $range): void
    {
        $this->service->spreadsheets_values->clear($spreadsheetId, $range, new \Google\Service\Sheets\ClearValuesRequest());
    }

    /**
     * Ensure a worksheet exists; create it if missing.
     * Returns the sheet title (same as $name).
     */
    public function ensureWorksheet(string $spreadsheetId, string $name, int $rows = 1000, int $cols = 2): string
    {
        // Check existing sheets
        $meta = $this->service->spreadsheets->get($spreadsheetId);
        foreach ($meta->getSheets() as $sheet) {
            if ($sheet->getProperties()->getTitle() === $name) {
                return $name; // Already exists
            }
        }

        // Create new sheet
        $addSheet = new AddSheetRequest([
            'properties' => new SheetProperties([
                'title'     => $name,
                'gridProperties' => [
                    'rowCount'    => $rows,
                    'columnCount' => $cols,
                ],
            ]),
        ]);
        $batchReq = new BatchUpdateSpreadsheetRequest([
            'requests' => [new Request(['addSheet' => $addSheet])],
        ]);
        $this->service->spreadsheets->batchUpdate($spreadsheetId, $batchReq);

        return $name;
    }

    // ──────────────────────────────────────────────────────
    //  Ticket map
    // ──────────────────────────────────────────────────────

    /**
     * Return a map of thread_id => row_number (1-indexed, skipping header).
     *
     * @return array<string, int>
     */
    public function getAllTickets(string $spreadsheetId, string $worksheet = 'Email log'): array
    {
        $rows = $this->getValues($spreadsheetId, "{$worksheet}!A:B");
        $map  = [];

        foreach ($rows as $idx => $row) {
            if ($idx === 0) {
                continue; // Skip header
            }
            $rowNum  = $idx + 1;       // 1-indexed
            $threadId = $row[1] ?? ''; // Column B
            if ($threadId !== '') {
                $map[$threadId] = $rowNum;
            }
        }

        return $map;
    }

    /**
     * Get a specific row's values.
     *
     * @return array<int, string>
     */
    public function getRow(string $spreadsheetId, string $worksheet, int $rowNumber): array
    {
        $range = "{$worksheet}!A{$rowNumber}:H{$rowNumber}";
        $values = $this->getValues($spreadsheetId, $range);
        return $values[0] ?? [];
    }

    // ──────────────────────────────────────────────────────
    //  Ticket ID counter
    // ──────────────────────────────────────────────────────

    /**
     * Read counter from Ticket_Config!B1, increment, save, and return "TCK-XXXXXX"
     */
    public function getNextTicketId(string $spreadsheetId): string
    {
        $raw  = $this->getCell($spreadsheetId, 'Ticket_Config!B1');
        $next = ((int) $raw) + 1;
        $this->setCell($spreadsheetId, 'Ticket_Config!B1', $next);
        return sprintf('TCK-%06d', $next);
    }

    // ──────────────────────────────────────────────────────
    //  Row builders
    // ──────────────────────────────────────────────────────

    /**
     * Build a ticket row ready for insertion / update.
     *
     * @return array<int, mixed>
     */
    public function createTicketRow(
        string $ticketId,
        string $threadId,
        string $fromEmail,
        string $subject,
        string $status,
        bool   $newSender = false
    ): array {
        $timestamp = date('Y-m-d H:i:s');
        $link      = sprintf(
            '=HYPERLINK("https://mail.google.com/mail/u/0/#inbox/%s","Open Mail")',
            $threadId
        );

        return [
            $ticketId,
            $threadId,
            $timestamp,
            $fromEmail,
            $subject,
            $status,
            $newSender ? 'Yes' : 'No',
            $link,
        ];
    }

    /**
     * Append a new ticket row to the worksheet.
     *
     * @param array<int, mixed> $rowData
     */
    public function addNewTicket(string $spreadsheetId, string $worksheet, array $rowData): void
    {
        $this->appendRow($spreadsheetId, "{$worksheet}!A:H", $rowData);
    }

    /**
     * Overwrite an existing ticket row.
     *
     * @param array<int, mixed> $rowData
     */
    public function updateExistingTicket(string $spreadsheetId, string $worksheet, int $rowNumber, array $rowData): void
    {
        $range = "{$worksheet}!A{$rowNumber}:H{$rowNumber}";
        $this->setValues($spreadsheetId, $range, [$rowData]);
    }

    // ──────────────────────────────────────────────────────
    //  Sync timestamp (Sync_State sheet)
    // ──────────────────────────────────────────────────────

    public function initializeStateSheets(string $spreadsheetId): void
    {
        // Sync_State
        $this->ensureWorksheet($spreadsheetId, 'Sync_State', 10, 2);
        $this->setValues($spreadsheetId, 'Sync_State!A1', [['Last Sync']]);

        // Thread_State
        $this->ensureWorksheet($spreadsheetId, 'Thread_State', 1000, 2);
        $this->setValues($spreadsheetId, 'Thread_State!A1', [['Thread ID', 'Last Processed Timestamp']]);
    }

    public function getLastSyncTimestamp(string $spreadsheetId): ?int
    {
        $value = $this->getCell($spreadsheetId, 'Sync_State!B1');
        return ($value !== null && $value !== '') ? (int) $value : null;
    }

    public function saveLastSyncTimestamp(string $spreadsheetId, int $timestamp): void
    {
        $this->ensureWorksheet($spreadsheetId, 'Sync_State', 10, 2);
        $this->setCell($spreadsheetId, 'Sync_State!B1', $timestamp);
    }

    // ──────────────────────────────────────────────────────
    //  Thread state (Thread_State sheet)
    // ──────────────────────────────────────────────────────

    /**
     * Load thread processing state from the Thread_State worksheet.
     *
     * @return array<string, int>   thread_id => timestamp
     */
    public function loadThreadStateFromSheet(string $spreadsheetId): array
    {
        $state = [];
        try {
            $rows = $this->getValues($spreadsheetId, 'Thread_State!A:B');
            foreach ($rows as $idx => $row) {
                if ($idx === 0 || count($row) < 2 || $row[0] === '') {
                    continue;
                }
                $state[$row[0]] = (int) $row[1];
            }
        } catch (\Exception) {
            // Sheet might not exist yet — that's fine
        }
        return $state;
    }

    /**
     * Persist thread state to the Thread_State worksheet.
     *
     * @param array<string, int> $state
     */
    public function saveThreadStateToSheet(string $spreadsheetId, array $state): void
    {
        $this->ensureWorksheet($spreadsheetId, 'Thread_State', 1000, 2);

        $data = [['Thread ID', 'Last Processed Timestamp']];
        foreach ($state as $tid => $ts) {
            $data[] = [$tid, $ts];
        }

        $this->clearRange($spreadsheetId, 'Thread_State!A:B');
        $this->setValues($spreadsheetId, 'Thread_State!A1', $data);
    }

   
    /**
     * @return string[]
     */


    // ──────────────────────────────────────────────────────
    //  Known senders
    // ──────────────────────────────────────────────────────

    /**
     * Read column D (index 3) from the Email log and return all unique emails.
     *
     * @return array<string, true>   email => true (set-like)
     */
    public function loadKnownSenders(string $spreadsheetId, string $worksheet = 'Email log'): array
    {
        $rows    = $this->getValues($spreadsheetId, "{$worksheet}!A:H");
        $senders = [];
        foreach ($rows as $idx => $row) {
            if ($idx === 0) {
                continue; // Skip header
            }
            $email = strtolower($row[3] ?? '');
            if ($email !== '') {
                $senders[$email] = true;
            }
        }
        return $senders;
    }

    // ──────────────────────────────────────────────────────
    //  All rows (for stale-ticket check)
    // ──────────────────────────────────────────────────────

    /**
     * @return array<int, array<int, string>>
     */
    public function getAllRows(string $spreadsheetId, string $worksheet = 'Email log'): array
    {
        return $this->getValues($spreadsheetId, "{$worksheet}!A:H");
    }

    /**
     * Update a single row in-place (used by auto-close).
     *
     * @param array<int, string> $row
     */
    public function updateRow(string $spreadsheetId, string $worksheet, int $rowNumber, array $row): void
    {
        $range = "{$worksheet}!A{$rowNumber}:H{$rowNumber}";
        $this->setValues($spreadsheetId, $range, [$row]);
    }

    /**
     * Delete a single row by index (used by auto-close with "delete" action).
     * NOTE: Deleting via the Sheets API requires a batchUpdate request.
     */
    public function deleteRow(string $spreadsheetId, int $sheetId, int $rowIndex): void
    {
        $request = new Request([
            'deleteDimension' => [
                'range' => [
                    'sheetId'    => $sheetId,
                    'dimension'  => 'ROWS',
                    'startIndex' => $rowIndex - 1, // 0-indexed
                    'endIndex'   => $rowIndex,
                ],
            ],
        ]);
        $batchReq = new BatchUpdateSpreadsheetRequest(['requests' => [$request]]);
        $this->service->spreadsheets->batchUpdate($spreadsheetId, $batchReq);
    }

    /**
     * Look up the numeric sheet ID for a given worksheet title.
     */
    public function getSheetId(string $spreadsheetId, string $worksheetTitle): ?int
    {
        $meta = $this->service->spreadsheets->get($spreadsheetId);
        foreach ($meta->getSheets() as $sheet) {
            if ($sheet->getProperties()->getTitle() === $worksheetTitle) {
                return $sheet->getProperties()->getSheetId();
            }
        }
        return null;
    }
}
