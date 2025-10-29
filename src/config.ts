export type Logger = {
  error?: (...args: any[]) => void
  warn?: (...args: any[]) => void
  info?: (...args: any[]) => void
  log?: (...args: any[]) => void
}

export type CredentialOptions = {
  sharedMailbox: string
  tenantId: string
  clientId: string
  clientSecret: string

  attachmentUploadOptions?: AttachmentUploadOptions
  logger?: Logger
}

export type AttachmentUploadOptions = {
  // 大文件的阈值，超过此大小将使用上传会话
  largeFileThreshold?: number
  // 分块上传的大小，必须是320KB的倍数
  chunkSize?: number
}
