# msgraph-shared-mailbox

![build-status](https://github.com/chakhsu/msgraph-shared-mailbox/actions/workflows/ci.yml/badge.svg) ![npm](https://img.shields.io/npm/v/msgraph-shared-mailbox) ![license](https://img.shields.io/npm/l/msgraph-shared-mailbox)

[English](./README.md) | [简体中文](./README_CN.md)

一个基于 Microsoft Graph 的 Node.js 共享邮箱发信库。支持文本或 HTML 正文；支持内存、磁盘路径或远程 URL 的附件；自动在小文件直传与大文件分块上传之间切换；并提供按邮件 ID 查询的能力。

## 特性

- 使用 Microsoft Graph 从共享邮箱发送邮件。
- 支持 `to`、`cc`、`bcc` 收件人；正文支持文本或 HTML。
- 附件支持三种来源：内存 `content`、本地路径 `path`、远程 URL `href`。
- 大附件自动启用上传会话，分块大小需为 320KB 的倍数。
- 收件人支持字符串或数组输入，自动去重，支持逗号/分号分隔。
- 通过邮件 ID 获取详情，可选是否同时获取附件列表。
- 简单的错误日志钩子。

## 安装

```bash
pnpm add msgraph-shared-mailbox
# 或
npm install msgraph-shared-mailbox
```

Node.js 18+。

## 前置条件

- 在 Azure AD 注册应用并创建客户端机密（Client Credentials 模式）。
- 授予并管理员同意以下 Microsoft Graph 应用权限：
  - 发送邮件：`Mail.Send`
  - 读取邮件/附件：`Mail.Read`（若需要读取）
- 准备共享邮箱地址，例如 `shared@microsoft.com`。

## 快速开始

```ts
import { MailClient } from 'msgraph-shared-mailbox'

const mailClient = new MailClient({
  tenantId: process.env.TENANT_ID!,
  clientId: process.env.CLIENT_ID!,
  clientSecret: process.env.CLIENT_SECRET!,
  sharedMailbox: process.env.SHARED_MAILBOX!,
  // 大附件上传设置（必填；以下为默认值）
  attachmentUploadOptions: {
    largeFileThreshold: 3 * 1024 * 1024, // 3MB
    chunkSize: 320 * 1024 // 320KB（必须是 320KB 的倍数）
  }
})

await mailClient.sendMail({
  subject: 'Hello from msgraph-shared-mailbox',
  to: 'someone@example.com', // 字符串（逗号/分号分隔）或数组
  text: '这是一段文本正文，你也可以使用 HTML。'
})
```

单例用法（Singleton）：

```ts
// 在应用启动时初始化一次
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

// 在其他地方直接复用
await mailClient.sendMail({ subject: '...', to: 'a@ex.com', text: '...' })
```

## 附件用法

可以通过以下方式提供附件：

- `content`：`Buffer | string`（适合小文件，内存直传）
- `path`：本地文件路径
- `href`：远程文件 URL（http/https）

示例：

```ts
await mailClient.sendMail({
  subject: '含附件的邮件',
  to: 'someone@example.com',
  text: '附件来自 content、path 与 href',
  attachments: [
    { filename: 'note.txt', content: Buffer.from('Hello attachment content') },
    { filename: 'sample.txt', path: '/absolute/path/to/sample.txt' },
    { filename: 'bytes.bin', href: 'https://httpbin.org/bytes/1024' }
  ]
})
```

限制：

- 单个附件大小必须 ≤ 150MB（Graph 上传会话限制）。
- 超过 `largeFileThreshold` 的附件会使用上传会话；`chunkSize` 必须是 320KB 的倍数。

## 按 ID 获取邮件

```ts
const message = await mailClient.getMailById('<MESSAGE_ID>', {
  includeAttachments: true
  // 如需自定义字段：
  // select: ['id', 'subject']
})
```

## API

### MailClient(options: CredentialOptions)

参数：

- `tenantId`：Azure AD 租户 ID。
- `clientId`：应用（客户端）ID。
- `clientSecret`：应用客户端机密。
- `sharedMailbox`：共享邮箱地址。
- `attachmentUploadOptions`：
  - `largeFileThreshold`：数字（默认 3MB）。
  - `chunkSize`：数字（默认 320KB，必须为 320KB 倍数）。
- `logger?`：可选，形如 `{ error?: (...args: any[]) => void }`。

### 单例模式说明

使用 `MailClient.getInstance(options)` 可创建进程级单例：

```ts
const mc = MailClient.getInstance({
  tenantId: '...',
  clientId: '...',
  clientSecret: '...',
  sharedMailbox: '...',
  attachmentUploadOptions: { largeFileThreshold: 3 * 1024 * 1024, chunkSize: 320 * 1024 }
})
```

注意：

- 只有第一次调用会使用提供的 `options` 进行初始化；后续调用会返回同一个实例，并忽略新的 `options`。
- 单租户/单共享邮箱的场景推荐使用单例。如果需要按请求切换不同的凭据或邮箱，请使用 `new MailClient(options)` 创建独立实例。

### sendMail(mail: MailOptions): Promise<string>

发送邮件并返回新邮件的 ID。

`MailOptions`：

- `subject`：字符串
- `to`：字符串或字符串数组（支持逗号/分号，自动去重）
- `cc?`：字符串或字符串数组
- `bcc?`：字符串或字符串数组
- 二选一：
  - `{ text: string }`
  - `{ html: string }`
- `attachments?`：`{ filename?, content?, path?, href? }[]`

### getMailById(id: string, options?: { select?: string[]; includeAttachments?: boolean }): Promise<any>

按 ID 查询邮件。`includeAttachments` 为 `true` 时将同时查询附件并合并到返回结果中。

## 运行示例

```bash
pnpm i

# 设置环境变量
export TENANT_ID=...
export CLIENT_ID=...
export CLIENT_SECRET=...
export SHARED_MAILBOX=shared@microsoft.com

# 基础发送
pnpm ts-node examples/basic-send.ts

# 带附件发送
pnpm ts-node examples/attachments.ts

# 通过 ID 查询邮件
export MESSAGE_ID=...
pnpm ts-node examples/get-mail-by-id.ts
```

## 常见问题

- 权限不足（`insufficient privileges`）：确保应用具有 `Mail.Send`（读取还需 `Mail.Read`）应用权限，并已获得管理员同意。
- 找不到共享邮箱：确认 `sharedMailbox` 地址正确且该邮箱在租户中存在。
- 大附件上传失败：确保 `chunkSize` 为 320KB 的倍数，且单附件不超过 150MB。
- 429 限流：考虑重试/退避策略；本库未实现自动重试。

## 许可

MIT © Chakhsu.Lau
