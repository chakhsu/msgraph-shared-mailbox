import fs from 'node:fs/promises'
import path from 'node:path'

import { ClientSecretCredential } from '@azure/identity'
import { Client } from '@microsoft/microsoft-graph-client'

import { getContentTypeFromFileName } from './utils'
import type { CredentialOptions, AttachmentUploadOptions, Logger } from './config'

const defaultAttachmentUploadOptions: Required<AttachmentUploadOptions> = {
  // 大文件的阈值，超过此大小将使用上传会话, 默认3MB
  largeFileThreshold: 3 * 1024 * 1024,
  // 分块上传的大小，必须是320KB的倍数, 默认320KB
  chunkSize: 320 * 1024
}

// 单个附件最大支持 150MB（Microsoft Graph 上传会话限制）
const MAX_ATTACHMENT_SIZE = 150 * 1024 * 1024

const defaultLogger: Required<Pick<Logger, 'error'>> = {
  error: (...args: any[]) => console.error(...args)
}

export type MailOptions = {
  subject: string
  to: string | string[]
  cc?: string | string[]
  bcc?: string | string[]
} & ({ text: string; html?: never } | { html: string; text?: never }) & {
    attachments?: {
      filename?: string
      content?: string | Buffer
      path?: string
      href?: string
    }[]
  }

export type getMailByIdOptions = {
  select: string[]
  includeAttachments?: boolean
  // 可选：明确指定传入的 id 类型
  // 'graph' 表示 Graph 消息 ID；'internetMessageId' 表示 RFC822 Internet Message ID（通常带尖括号）
  idType?: 'graph' | 'internetMessageId'
}

export class MailClient {
  private static _instance: MailClient | null = null
  private sharedMailbox: string
  private credential: ClientSecretCredential
  private attachmentUploadOptions: Required<AttachmentUploadOptions>
  private client: Client
  private logger: Required<Pick<Logger, 'error'>>

  constructor(options: CredentialOptions) {
    this.sharedMailbox = options.sharedMailbox
    this.credential = new ClientSecretCredential(options.tenantId, options.clientId, options.clientSecret)

    this.attachmentUploadOptions = {
      ...defaultAttachmentUploadOptions,
      ...options.attachmentUploadOptions
    }
    // 规范化分块大小为 320KB 倍数
    {
      const base = 320 * 1024
      let chunkSize = this.attachmentUploadOptions.chunkSize
      if (chunkSize % base !== 0) {
        chunkSize = Math.max(base, Math.floor(chunkSize / base) * base)
      }
      this.attachmentUploadOptions.chunkSize = chunkSize
    }

    this.client = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          const token = await this.credential.getToken('https://graph.microsoft.com/.default')
          return token.token
        }
      }
    })

    this.logger = { ...defaultLogger, ...(options.logger || {}) }
  }

  static getInstance(options: CredentialOptions): MailClient {
    if (!MailClient._instance) {
      MailClient._instance = new MailClient(options)
    }
    return MailClient._instance
  }

  async sendMail(mail: MailOptions) {
    try {
      const toRecipients = this.buildRecipients(mail.to)
      if (!toRecipients || toRecipients.length === 0) {
        throw new Error('No recipients provided in "to"')
      }
      const ccRecipients = this.buildRecipients(mail.cc)
      const bccRecipients = this.buildRecipients(mail.bcc)

      let draftMessage: any = {
        subject: mail.subject,
        toRecipients,
        ccRecipients,
        bccRecipients
      }
      if (mail.html) {
        draftMessage.body = {
          contentType: 'html',
          content: mail.html
        }
      } else if (mail.text) {
        draftMessage.body = {
          contentType: 'text',
          content: mail.text
        }
      }

      // 创建草稿邮件
      const draft = await this.client.api(`/users/${this.sharedMailbox}/messages`).post(draftMessage)

      await this.uploadAttachments(draft.id, mail.attachments)

      // 正式发送邮件
      await this.client.api(`/users/${this.sharedMailbox}/messages/${draft.id}/send`).post({})

      return {
        id: draft.id,
        internetMessageId: draft.internetMessageId
      }
    } catch (error) {
      this.logger.error('Failed to send mail', error)
      throw error
    }
  }

  private buildRecipients(input: string | string[] | undefined): { emailAddress: { address: string } }[] | undefined {
    if (!input) return undefined
    const parts = Array.isArray(input) ? input : [input]
    const addresses: string[] = []
    for (const p of parts) {
      const split = p.split(/[;,]+/)
      for (const s of split) {
        const addr = s.trim()
        if (addr) addresses.push(addr)
      }
    }
    const unique = Array.from(new Set(addresses))
    return unique.length ? unique.map((address) => ({ emailAddress: { address } })) : undefined
  }

  private async uploadAttachments(messageId: string, attachments: MailOptions['attachments'] | undefined) {
    if (!attachments || attachments.length === 0) {
      return
    }

    // 使用已在构造函数中归一化的阈值

    for await (const attachment of attachments) {
      if (!attachment.content && !attachment.path && !attachment.href) {
        continue
      }

      let fileName = attachment.filename
      if (!fileName) {
        if (attachment.path) {
          fileName = attachment.path.split('/').pop() || 'attachment'
        } else if (attachment.href) {
          try {
            const u = new URL(attachment.href)
            fileName = path.basename(u.pathname) || 'attachment'
          } catch {
            fileName = attachment.href?.split('/').pop() || 'attachment'
          }
        } else {
          fileName = 'attachment'
        }
      }
      let contentType = getContentTypeFromFileName(fileName)

      // 如果提供了内存内容，先检查大小，再根据大小选择普通上传或分块上传
      if (attachment.content) {
        const buf = Buffer.isBuffer(attachment.content) ? attachment.content : Buffer.from(attachment.content)
        if (buf.length > MAX_ATTACHMENT_SIZE) {
          throw new Error('Attachment size exceeds maximum 150MB')
        }
        if (buf.length <= this.attachmentUploadOptions.largeFileThreshold) {
          const base64Content = buf.toString('base64')
          await this.client.api(`/users/${this.sharedMailbox}/messages/${messageId}/attachments`).post({
            '@odata.type': '#microsoft.graph.fileAttachment',
            name: fileName,
            contentBytes: base64Content,
            contentType: contentType
          })
        } else {
          await this.uploadLargeContent(messageId, fileName, buf, contentType)
        }
        continue
      }

      if (attachment.path) {
        // 判断是否为本地文件 和文件大小
        let stats: any
        try {
          stats = await fs.stat(attachment.path)
        } catch {
          // 文件不存在或不可访问
          continue
        }
        if (!stats.isFile() || stats.size === 0) {
          continue
        }

        if (stats.size > MAX_ATTACHMENT_SIZE) {
          throw new Error('Attachment size exceeds maximum 150MB')
        }

        // 如果是小文件，直接上传
        if (stats.size <= this.attachmentUploadOptions.largeFileThreshold) {
          const fileContent = await fs.readFile(attachment.path)
          let base64Content = Buffer.from(fileContent).toString('base64')

          await this.client.api(`/users/${this.sharedMailbox}/messages/${messageId}/attachments`).post({
            '@odata.type': '#microsoft.graph.fileAttachment',
            name: fileName,
            contentBytes: base64Content,
            contentType: getContentTypeFromFileName(fileName)
          })
        } else {
          // 如果是大文件，使用上传会话
          await this.uploadLargeFileByPath(messageId, fileName, attachment.path, contentType, stats.size)
        }

        continue
      }

      if (attachment.href) {
        await this.uploadAttachmentFromUrl(messageId, fileName, attachment.href, contentType)
        continue
      }
    }
  }

  private async uploadAttachmentFromUrl(messageId: string, fileName: string, href: string, contentType: string) {
    // 仅处理 http/https
    if (!/^https?:\/\//i.test(href)) {
      return
    }

    // 先尝试 HEAD 拿到大小
    let size: number | undefined
    try {
      const head = await fetch(href, { method: 'HEAD' })
      if (head.ok) {
        const cl = head.headers.get('content-length')
        if (cl) {
          const parsed = Number(cl)
          if (!Number.isNaN(parsed) && parsed >= 0) {
            size = parsed
          }
        }
      }
    } catch {}

    if (size !== undefined && size > MAX_ATTACHMENT_SIZE) {
      throw new Error('Attachment size exceeds maximum 150MB')
    }
    if (size !== undefined && size <= this.attachmentUploadOptions.largeFileThreshold) {
      // 小文件直接下载到内存并作为普通附件上传
      const res = await fetch(href)
      if (!res.ok) return
      const arrayBuf = await res.arrayBuffer()
      const buf = Buffer.from(arrayBuf)
      if (buf.length > MAX_ATTACHMENT_SIZE) {
        throw new Error('Attachment size exceeds maximum 150MB')
      }
      const base64Content = buf.toString('base64')
      await this.client.api(`/users/${this.sharedMailbox}/messages/${messageId}/attachments`).post({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: fileName,
        contentBytes: base64Content,
        contentType: contentType
      })
      return
    }

    // 直接使用 fetch 流分块上传（若 Content-Length 已知则边下边传；未知则暂存内存后上传）
    const res = await fetch(href)
    if (!res.ok || !res.body) return
    const contentLengthHeader = res.headers.get('content-length')
    if (size === undefined && contentLengthHeader) {
      const parsed = Number(contentLengthHeader)
      if (!Number.isNaN(parsed) && parsed >= 0) size = parsed
    }

    if (size !== undefined && size > MAX_ATTACHMENT_SIZE) {
      throw new Error('Attachment size exceeds maximum 150MB')
    }

    // 规范化分块大小为 320KB 的倍数
    const base = 320 * 1024
    let chunkSize = this.attachmentUploadOptions.chunkSize
    if (chunkSize % base !== 0) {
      chunkSize = Math.max(base, Math.floor(chunkSize / base) * base)
    }

    // 如果已知总大小，创建会话并边下边传
    if (size !== undefined && size > this.attachmentUploadOptions.largeFileThreshold) {
      const sessionBody: any = {
        AttachmentItem: {
          attachmentType: 'file',
          name: fileName,
          size: size,
          contentType: contentType
        }
      }
      const session = await this.client.api(`/users/${this.sharedMailbox}/messages/${messageId}/attachments/createUploadSession`).post(sessionBody)
      const uploadUrl = session.uploadUrl

      const stream = res.body as any
      const reader = stream.getReader()
      let offset = 0
      let pending = Buffer.alloc(0)
      while (true) {
        const { done, value } = await reader.read()
        if (done) break
        const chunkBuf = Buffer.from(value)
        pending = pending.length === 0 ? chunkBuf : Buffer.concat([pending, chunkBuf])
        while (pending.length >= chunkSize) {
          const toSend = pending.subarray(0, chunkSize)
          const start = offset
          const end = start + toSend.length - 1
          const resp = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
              'Content-Length': `${toSend.length}`,
              'Content-Range': `bytes ${start}-${end}/${size}`,
              'Content-Type': 'application/octet-stream'
            },
            body: toSend
          })
          if (!resp.ok && resp.status !== 202 && resp.status !== 201 && resp.status !== 200) {
            throw new Error(`Upload chunk failed with status ${resp.status}`)
          }
          offset += toSend.length
          pending = pending.subarray(toSend.length)
        }
      }
      // 发送剩余部分作为最后一块
      if (pending.length > 0) {
        const start = offset
        const end = size - 1
        const resp = await fetch(uploadUrl, {
          method: 'PUT',
          headers: {
            'Content-Length': `${pending.length}`,
            'Content-Range': `bytes ${start}-${end}/${size}`,
            'Content-Type': 'application/octet-stream'
          },
          body: pending
        })
        if (!resp.ok && resp.status !== 202 && resp.status !== 201 && resp.status !== 200) {
          throw new Error(`Upload final chunk failed with status ${resp.status}`)
        }
      }
      return
    }

    // 未知总大小或小于阈值：下载至内存后按大小选择上传方式（不落地文件）
    const stream = res.body as any
    const reader = stream.getReader()
    const buffers: Buffer[] = []
    let total = 0
    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      const buf = Buffer.from(value)
      buffers.push(buf)
      total += buf.length
      if (total > MAX_ATTACHMENT_SIZE) {
        throw new Error('Attachment size exceeds maximum 150MB')
      }
    }
    const all = Buffer.concat(buffers)
    if (total <= this.attachmentUploadOptions.largeFileThreshold) {
      const base64Content = all.toString('base64')
      await this.client.api(`/users/${this.sharedMailbox}/messages/${messageId}/attachments`).post({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: fileName,
        contentBytes: base64Content,
        contentType: contentType
      })
    } else {
      await this.uploadLargeContent(messageId, fileName, all, contentType)
    }
  }

  private async uploadLargeContent(messageId: string, fileName: string, content: Buffer, contentType: string) {
    // 规范化分块大小为 320KB 的倍数
    const base = 320 * 1024
    let chunkSize = this.attachmentUploadOptions.chunkSize
    if (chunkSize % base !== 0) {
      chunkSize = Math.max(base, Math.floor(chunkSize / base) * base)
    }

    if (content.length > MAX_ATTACHMENT_SIZE) {
      throw new Error('Attachment size exceeds maximum 150MB')
    }

    const session = await this.client.api(`/users/${this.sharedMailbox}/messages/${messageId}/attachments/createUploadSession`).post({
      AttachmentItem: {
        attachmentType: 'file',
        name: fileName,
        size: content.length,
        contentType: contentType
      }
    })

    const uploadUrl = session.uploadUrl
    let offset = 0
    while (offset < content.length) {
      const endExclusive = Math.min(offset + chunkSize, content.length)
      const chunk = content.subarray(offset, endExclusive)
      const start = offset
      const end = endExclusive - 1
      const res = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': `${chunk.length}`,
          'Content-Range': `bytes ${start}-${end}/${content.length}`,
          'Content-Type': 'application/octet-stream'
        },
        body: chunk
      })
      if (!res.ok && res.status !== 202 && res.status !== 201 && res.status !== 200) {
        throw new Error(`Upload chunk failed with status ${res.status}`)
      }
      offset = endExclusive
    }
  }

  private async uploadLargeFileByPath(messageId: string, fileName: string, filePath: string, contentType: string, fileSize: number) {
    // 规范化分块大小为 320KB 的倍数
    const base = 320 * 1024
    let chunkSize = this.attachmentUploadOptions.chunkSize
    if (chunkSize % base !== 0) {
      chunkSize = Math.max(base, Math.floor(chunkSize / base) * base)
    }

    if (fileSize > MAX_ATTACHMENT_SIZE) {
      throw new Error('Attachment size exceeds maximum 150MB')
    }

    const session = await this.client.api(`/users/${this.sharedMailbox}/messages/${messageId}/attachments/createUploadSession`).post({
      AttachmentItem: {
        attachmentType: 'file',
        name: fileName,
        size: fileSize,
        contentType: contentType
      }
    })

    const uploadUrl = session.uploadUrl

    const fd = await fs.open(filePath, 'r')
    const fileStream = fd.createReadStream({ highWaterMark: chunkSize })

    let offset = 0
    for await (const chunk of fileStream) {
      const chunkData = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk as any)
      const start = offset
      const end = start + chunkData.length - 1
      const res = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': `${chunkData.length}`,
          'Content-Range': `bytes ${start}-${end}/${fileSize}`,
          'Content-Type': 'application/octet-stream'
        },
        body: chunkData
      })
      // 中间分块通常返回 202，最后一块返回 201
      if (!res.ok && res.status !== 202 && res.status !== 201 && res.status !== 200) {
        throw new Error(`Upload chunk failed with status ${res.status}`)
      }
      offset += chunkData.length
    }
    // 关闭流与文件句柄
    fileStream.close()
    await fd.close()
  }

  // 根据邮件ID查询邮件（支持 Graph ID 或 Internet Message ID）
  async getMailById(id: string, options: Partial<getMailByIdOptions> = {}) {
    const { select, includeAttachments = false, idType } = options

    if (!id) {
      throw new Error('id is required')
    }

    const selectFields =
      Array.isArray(select) && select.length > 0 ? select.join(',') : 'id,subject,bodyPreview,body,sentDateTime,receivedDateTime,from,toRecipients,ccRecipients,bccRecipients,hasAttachments'

    // 判断是否为 Internet Message ID（通常形如 <...@...>）或显式指定 idType
    const looksLikeInternetId = /^<[^>]+@[^>]+>$/.test(id)
    const useInternetMessageId = idType === 'internetMessageId' || looksLikeInternetId

    if (useInternetMessageId) {
      // 通过 OData $filter 使用 internetMessageId 精确匹配
      const filterValue = id.replace(/'/g, "''")
      const list: any = await this.client.api(`/users/${this.sharedMailbox}/messages`).filter(`internetMessageId eq '${filterValue}'`).select(selectFields).get()

      const message = list?.value?.[0]
      if (!message) {
        throw new Error('Message not found by internetMessageId')
      }

      if (includeAttachments) {
        const attRes = await this.client.api(`/users/${this.sharedMailbox}/messages/${message.id}/attachments`).get()
        message.attachments = attRes.value || []
      }

      return message
    }

    // 默认按 Graph 消息 ID 查询
    const message = await this.client.api(`/users/${this.sharedMailbox}/messages/${id}`).select(selectFields).get()

    if (includeAttachments) {
      const attRes = await this.client.api(`/users/${this.sharedMailbox}/messages/${id}/attachments`).get()
      message.attachments = attRes.value || []
    }

    return message
  }
}
