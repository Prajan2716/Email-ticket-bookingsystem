<?php
/**
 * Gmail Ticket System — Web Entry Point
 * PHP equivalent of app.py (Flask + APScheduler)
 *
 * Run with the built-in server:
 *   php -S 0.0.0.0:8080 public/index.php
 *
 * Or with any PHP-FPM / Apache / Nginx setup.
 *
 * Background sync is handled by a cron job running sync.php every minute,
 * NOT by this file.  This file only shows status and triggers manual syncs.
 */

declare(strict_types=1);

require_once __DIR__ . '/../vendor/autoload.php';

use GmailTicket\TicketSyncer;

// ── Simple router ────────────────────────────────────────────────────────────

$path = parse_url($_SERVER['REQUEST_URI'] ?? '/', PHP_URL_PATH);

match ($path) {
    '/status'       => handleStatus(),
    '/sync'         => handleManualSync(),
    default         => handleHome(),
};

// ── Handlers ─────────────────────────────────────────────────────────────────

function handleHome(): void
{
    $statusFile = __DIR__ . '/../sync_status.json';
    $status     = file_exists($statusFile)
        ? json_decode(file_get_contents($statusFile), true)
        : ['last_sync' => 'Never', 'last_error' => null, 'sync_count' => 0];

    header('Content-Type: text/html; charset=utf-8');
    echo <<<HTML
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="refresh" content="10">
        <title>Gmail Ticket System</title>
        <style>
            body      { font-family: Arial, sans-serif; margin: 40px; background: #fafafa; }
            h1        { color: #333; }
            .card     { background: #fff; border: 1px solid #ddd; border-radius: 8px;
                        padding: 24px; max-width: 540px; box-shadow: 0 2px 6px rgba(0,0,0,.06); }
            .ok       { color: #1a7f37; font-weight: bold; }
            .err      { color: #cf222e; }
            .meta     { color: #555; font-size: .9em; margin: 6px 0; }
            a.btn     { display: inline-block; margin-top: 16px; padding: 10px 20px;
                        background: #0969da; color: #fff; border-radius: 6px;
                        text-decoration: none; font-size: .95em; }
            a.btn:hover { background: #0757ba; }
        </style>
    </head>
    <body>
        <h1>📧 Gmail Ticket System</h1>
        <div class="card">
            <p class="ok">✅ System is running</p>
            <p class="meta">⏱  Sync runs every minute via cron</p>
            <p class="meta">📊 Total syncs: {$status['sync_count']}</p>
            <p class="meta">🕐 Last sync: {$status['last_sync']}</p>
    HTML;

    if ($status['last_error']) {
        echo "<p class=\"err\">❌ Last error: " . htmlspecialchars((string) $status['last_error']) . "</p>\n";
    }

    echo <<<HTML
            <a class="btn" href="/status">View JSON status</a>
            <a class="btn" href="/sync" style="background:#1a7f37;margin-left:8px">▶ Run sync now</a>
        </div>
        <p style="color:#888;font-size:.8em;margin-top:16px">Page auto-refreshes every 10 s</p>
    </body>
    </html>
    HTML;
}

function handleStatus(): void
{
    $statusFile = __DIR__ . '/../sync_status.json';
    $status     = file_exists($statusFile)
        ? json_decode(file_get_contents($statusFile), true)
        : ['last_sync' => null, 'last_error' => null, 'sync_count' => 0];

    header('Content-Type: application/json');
    echo json_encode([
        'scheduler_running' => true, // cron is always "running"
        'last_sync'         => $status['last_sync'],
        'last_error'        => $status['last_error'],
        'total_syncs'       => $status['sync_count'],
    ], JSON_PRETTY_PRINT);
}

function handleManualSync(): void
{
    $startTime = microtime(true);
    $error     = null;

    try {
        $syncer = new TicketSyncer();
        $syncer->syncMailToSheet();
        updateSyncStatus(null);
    } catch (\Throwable $e) {
        $error = $e->getMessage();
        updateSyncStatus($error);
    }

    $elapsed = round(microtime(true) - $startTime, 2);

    header('Content-Type: application/json');
    echo json_encode([
        'success' => $error === null,
        'message' => $error ?? 'Sync completed successfully',
        'elapsed_seconds' => $elapsed,
    ], JSON_PRETTY_PRINT);
}

function updateSyncStatus(?string $error): void
{
    $statusFile = __DIR__ . '/../sync_status.json';
    $current    = file_exists($statusFile)
        ? json_decode(file_get_contents($statusFile), true)
        : ['sync_count' => 0];

    $current['last_sync']  = date('Y-m-d H:i:s');
    $current['last_error'] = $error;
    $current['sync_count'] = ($current['sync_count'] ?? 0) + 1;

    file_put_contents($statusFile, json_encode($current, JSON_PRETTY_PRINT));
}
