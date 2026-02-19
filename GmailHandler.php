<?php
/**
 * GmailHandler - Manages all Gmail API operations
 */

declare(strict_types=1);

namespace GmailTicket;

use Google\Client;
use Google\Service\Gmail;

class GmailHandler
{
    private const SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/gmail.readonly',
        'https://www.googleapis.com/auth/gmail.modify',
    ];

    private const TOKEN_FILE    = __DIR__ . '/../token.json';
    private const CREDS_FILE    = __DIR__ . '/../credentials.json';

    private const NOREPLY_PATTERNS = [
        'noreply@',
        'no-reply@',
        'no_reply@',
        'donotreply@',
        'do-not-reply@',
        'do_not_reply@',
        'notifications@',
        'notification@',
        'automated@',
        'automation@',
        'mailer@',
        'daemon@',
        'bounce@',
        'bounces@',
    ];

    // ──────────────────────────────────────────────────────
    //  Authentication
    // ──────────────────────────────────────────────────────

    public function getClient(): Client
    {
        $client = new Client();
        $client->setAuthConfig(self::CREDS_FILE);
        $client->setScopes(self::SCOPES);
        $client->setAccessType('offline');
        $client->setPrompt('select_account consent');

        if (file_exists(self::TOKEN_FILE)) {
            $token = json_decode(file_get_contents(self::TOKEN_FILE), true);
            $client->setAccessToken($token);
        }

        // Refresh token if expired
        if ($client->isAccessTokenExpired()) {
            if ($client->getRefreshToken()) {
                $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
                file_put_contents(self::TOKEN_FILE, json_encode($client->getAccessToken()));
            } else {
                // No refresh token — run interactive OAuth flow
                $authUrl = $client->createAuthUrl();
                printf("Open the following link in your browser:\n%s\n\n", $authUrl);
                print 'Enter verification code: ';
                $authCode = trim((string) fgets(STDIN));
                $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);
                $client->setAccessToken($accessToken);

                if (array_key_exists('error', $accessToken)) {
                    throw new \RuntimeException(implode(', ', $accessToken));
                }

                file_put_contents(self::TOKEN_FILE, json_encode($client->getAccessToken()));
            }
        }

        return $client;
    }

    public function getService(): Gmail
    {
        return new Gmail($this->getClient());
    }

    // ──────────────────────────────────────────────────────
    //  User helpers
    // ──────────────────────────────────────────────────────

    public function getMyEmail(Gmail $gmail): string
    {
        $profile = $gmail->users->getProfile('me');
        return strtolower($profile->getEmailAddress());
    }

    // ──────────────────────────────────────────────────────
    //  Email address utilities
    // ──────────────────────────────────────────────────────

    /**
     * Extract bare email from strings like "Name <email@example.com>"
     */
    public function extractEmail(string $value): string
    {
        if (preg_match('/<(.+?)>/', $value, $m)) {
            return strtolower($m[1]);
        }
        return strtolower(trim($value));
    }

    /**
     * Return true if the address is a no-reply / automated sender
     */
    public function isNoreplyEmail(string $email): bool
    {
        $email = strtolower($email);
        foreach (self::NOREPLY_PATTERNS as $pattern) {
            if (str_starts_with($email, $pattern)) {
                return true;
            }
        }
        return false;
    }

    // ──────────────────────────────────────────────────────
    //  Labels
    // ──────────────────────────────────────────────────────

    public function getOrCreateLabel(Gmail $gmail, string $name): string
    {
        $result = $gmail->users_labels->listUsersLabels('me');
        foreach ($result->getLabels() as $label) {
            if ($label->getName() === $name) {
                return $label->getId();
            }
        }

        // Create new label
        $labelBody = new Gmail\Label([
            'name'                  => $name,
            'labelListVisibility'   => 'labelShow',
            'messageListVisibility' => 'show',
        ]);
        $created = $gmail->users_labels->create('me', $labelBody);
        return $created->getId();
    }

    // ──────────────────────────────────────────────────────
    //  Thread retrieval
    // ──────────────────────────────────────────────────────

    /**
     * Fetch ALL threads matching $query (handles pagination automatically)
     *
     * @return array<int, array{id: string, historyId: string}>
     */
    public function fetchAllThreads(Gmail $gmail, string $query): array
    {
        $threads   = [];
        $pageToken = null;

        do {
            $params = ['q' => $query, 'maxResults' => 100];
            if ($pageToken) {
                $params['pageToken'] = $pageToken;
            }

            $response  = $gmail->users_threads->listUsersThreads('me', $params);
            $batch     = $response->getThreads() ?? [];

            foreach ($batch as $t) {
                $threads[] = ['id' => $t->getId(), 'historyId' => $t->getHistoryId()];
            }

            $pageToken = $response->getNextPageToken();
        } while ($pageToken);

        return $threads;
    }

    /**
     * Get full thread detail including all messages
     */
    public function getThreadDetails(Gmail $gmail, string $threadId): Gmail\Thread
    {
        return $gmail->users_threads->get('me', $threadId);
    }

    /**
     * Return [message, headers] for the most-recent message in a thread.
     * Returns [null, null] if no messages exist.
     *
     * @return array{0: Gmail\Message|null, 1: array<string,string>|null}
     */
    public function getLastMessage(Gmail\Thread $thread): array
    {
        $messages = $thread->getMessages();
        if (!$messages) {
            return [null, null];
        }

        $last = null;
        foreach ($messages as $msg) {
            if ($last === null || (int) $msg->getInternalDate() > (int) $last->getInternalDate()) {
                $last = $msg;
            }
        }

        /** @var Gmail\Message $last */
        $rawHeaders = $last->getPayload()->getHeaders();
        $headers    = [];
        foreach ($rawHeaders as $h) {
            $headers[$h->getName()] = $h->getValue();
        }

        return [$last, $headers];
    }

    // ──────────────────────────────────────────────────────
    //  Label management on threads
    // ──────────────────────────────────────────────────────

    /**
     * @param string[]|null $addLabels
     * @param string[]|null $removeLabels
     */
    public function updateThreadLabels(
        Gmail $gmail,
        string $threadId,
        ?array $addLabels    = null,
        ?array $removeLabels = null
    ): void {
        $body = new Gmail\ModifyThreadRequest();

        if ($addLabels) {
            $body->setAddLabelIds($addLabels);
        }
        if ($removeLabels) {
            $body->setRemoveLabelIds($removeLabels);
        }

        $gmail->users_threads->modify('me', $threadId, $body);
    }

    /**
     * Move a thread to Trash
     */
    public function trashThread(Gmail $gmail, string $threadId): void
    {
        $gmail->users_threads->trash('me', $threadId);
    }
}
