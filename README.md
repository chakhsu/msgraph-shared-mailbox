# msgraph-shared-mailbox

![build-status](https://github.com/chakhsu/msgraph-shared-mailbox/actions/workflows/ci.yml/badge.svg) ![npm](https://img.shields.io/npm/v/msgraph-shared-mailbox) ![license](https://img.shields.io/npm/l/msgraph-shared-mailbox)

[English](./README.md) | [简体中文](./README_CN.md)

Email library for Node.js that sends mail from an Exchange Online shared mailbox via Microsoft Graph. It supports text or HTML bodies, small and large attachments (buffer, local file path, or remote URL), and fetching mail by ID.

## Features

- Send email from a shared mailbox using Microsoft Graph.
- Text or HTML body with `to`, `cc`, `bcc` recipients.
- Attachments from memory (`content`), local path (`path`), or URL (`href`).
- Automatic large file upload via Graph Upload Session when over threshold.
- Recipient parsing for strings or arrays; deduplicates comma/semicolon-separated addresses.
- Fetch message by ID with optional attachments.
- Minimal logging hook for error reporting.

## Installation

```bash
pnpm add msgraph-shared-mailbox
# or
npm install msgraph-shared-mailbox
```

Requires Node.js 18+.

## Prerequisites

- Azure AD App Registration with a Client Secret (Client Credentials flow).
- Application permissions granted and consented:
  - `Mail.Send` for sending emails.
  - `Mail.Read` if you will read messages or attachments.
- A valid shared mailbox address in your tenant, e.g. `shared@microsoft.com`.

## Quickstart

```ts
import { MailClient } from 'msgraph-shared-mailbox'

const mailClient = new MailClient({
  tenantId: process.env.TENANT_ID!,
  clientId: process.env.CLIENT_ID!,
  clientSecret: process.env.CLIENT_SECRET!,
  sharedMailbox: process.env.SHARED_MAILBOX!,
  // Large upload settings (defaults shown)
  attachmentUploadOptions: {
    largeFileThreshold: 3 * 1024 * 1024, // 3MB
    chunkSize: 320 * 1024 // 320KB (must be 320KB multiple)
  }
})

await mailClient.sendMail({
  subject: 'Hello from msgraph-shared-mailbox',
  to: 'someone@example.com', // string with comma/semicolon or an array
  text: 'This is a text body. You can also use HTML.'
})
```

Singleton alternative:

```ts
// Initialize once at app startup
import { MailClient } from 'msgraph-shared-mailbox'

export const mailClient = MailClient.getInstance({
  tenantId: process.env.TENANT_ID!,
  clientId: process.env.CLIENT_ID!,
  clientSecret: process.env.CLIENT_SECRET!,
  sharedMailbox: process.env.SHARED_MAILBOX!,
  attachmentUploadOptions: {
    largeFileThreshold: 3 * 1024 * 1024,
    chunkSize: 320 * 1024
  }
})

// Use anywhere else
await mailClient.sendMail({ subject: '...', to: 'a@ex.com', text: '...' })
```

## Attachments

Provide attachments via any of the following fields:

- `content`: `Buffer | string` kept in memory (small files)
- `path`: local file path
- `href`: remote file URL (http/https)

Example:

```ts
await mailClient.sendMail({
  subject: 'Email with attachments',
  to: 'someone@example.com',
  text: 'Attachments from content, path and href',
  attachments: [
    { filename: 'note.txt', content: Buffer.from('Hello attachment content') },
    { filename: 'sample.txt', path: '/absolute/path/to/sample.txt' },
    { filename: 'bytes.bin', href: 'https://httpbin.org/bytes/1024' }
  ]
})
```

Limits:

- Single attachment size must be ≤ 150MB (Graph upload session limit).
- Attachments larger than `largeFileThreshold` use an upload session with `chunkSize` that must be a multiple of 320KB.

## Fetch a Message by ID

```ts
const message = await mailClient.getMailById('<MESSAGE_ID>', {
  includeAttachments: true
  // choose fields if needed
  // select: ['id', 'subject']
})
```

You can also fetch by the RFC822 Internet Message ID (`internetMessageId`), which typically includes angle brackets like `'<...@...>'`:

```ts
// Option A: auto-detection when the ID looks like <...@...>
const msg1 = await mailClient.getMailById('<xxx@xxx.prod.exchangelabs.com>', {
  includeAttachments: true
})

// Option B: explicitly specify idType
const msg2 = await mailClient.getMailById('<xxx@xxx.prod.exchangelabs.com>', {
  idType: 'internetMessageId',
  includeAttachments: true,
  select: ['id', 'subject']
})
```

## API

### MailClient(options: CredentialOptions)

Options:

- `tenantId`: Azure AD tenant ID.
- `clientId`: App registration (client) ID.
- `clientSecret`: Client secret for the app.
- `sharedMailbox`: Shared mailbox email address.
- `attachmentUploadOptions?`: (optional) Large upload settings.
  - `largeFileThreshold?`: number (default 3MB).
  - `chunkSize?`: number (default 320KB, must be 320KB multiple).
- `logger?`: `{ error?: (...args: any[]) => void }` (optional).

### Singleton Mode

Use `MailClient.getInstance(options)` to create a process-wide singleton:

```ts
const mc = MailClient.getInstance({
  tenantId: '...',
  clientId: '...',
  clientSecret: '...',
  sharedMailbox: '...',
  attachmentUploadOptions: { largeFileThreshold: 3 * 1024 * 1024, chunkSize: 320 * 1024 }
})
```

Notes:

- Only the first call initializes the instance; later calls return the same instance and ignore new options.
- Prefer singleton for single-tenant apps. If you need different credentials/mailboxes per request, create separate `new MailClient(options)` instances instead.

### sendMail(mail: MailOptions): Promise<string>

Sends and returns the new message ID.

Return value details:

- Returns the Graph message `id` as a `string` and `internetMessageId` as a `string`.
- Implementation: creates a draft (`POST /users/{mailbox}/messages`), uploads attachments, then sends (`POST /users/{mailbox}/messages/{id}/send`). The returned `id` is the sent message ID.
- Use this `id` with `getMailById(id, { includeAttachments: true })` to retrieve message details or attachments.
- If you need the RFC822 `internetMessageId`, query the message and read `internetMessageId`, or later fetch by it using `getMailById('<...@...>', { idType: 'internetMessageId' })`.

`MailOptions`:

- `subject`: string
- `to`: string | string[] (supports comma/semicolon; deduplicated)
- `cc?`: string | string[]
- `bcc?`: string | string[]
- One of:
  - `{ text: string }`
  - `{ html: string }`
- `attachments?`: Array of `{ filename?, content?, path?, href? }`

### getMailById(id: string, options?: { select?: string[]; includeAttachments?: boolean }): Promise<any>

Fetches a message by ID. When `includeAttachments` is true, attachments are queried and merged into the result.

## Run the Examples

```bash
pnpm i

# Set required environment variables
export TENANT_ID=...
export CLIENT_ID=...
export CLIENT_SECRET=...
export SHARED_MAILBOX=shared@microsoft.com

# Basic send
pnpm ts-node examples/basic-send.ts

# Send with attachments
pnpm ts-node examples/attachments.ts

# Fetch a message by ID
export MESSAGE_ID=...
pnpm ts-node examples/get-mail-by-id.ts
```

## Troubleshooting

- `insufficient privileges`: ensure the app has `Mail.Send` (and `Mail.Read` if reading) application permissions with admin consent.
- Shared mailbox not found: verify `sharedMailbox` address and that the mailbox exists in your tenant.
- Large attachments failing: set `chunkSize` to a multiple of 320KB and keep each attachment ≤ 150MB.
- Rate limits (HTTP 429): consider retry/backoff; this library does not implement automatic retries.

## License

MIT © Chakhsu.Lau
