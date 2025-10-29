import { MailClient } from '../src'

const TENANT_ID = process.env.TENANT_ID || ''
const CLIENT_ID = process.env.CLIENT_ID || ''
const CLIENT_SECRET = process.env.CLIENT_SECRET || ''
const SHARED_MAILBOX = process.env.SHARED_MAILBOX || ''

async function main() {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !SHARED_MAILBOX) {
    throw new Error('Please set TENANT_ID, CLIENT_ID, CLIENT_SECRET, SHARED_MAILBOX environment variables')
  }

  const logger = {
    error: (...args: any[]) => {
      // 替换为你的日志框架（例如 winston/pino）
      console.log('[MailClient Error]', ...args)
    }
  }

  const mailClient = new MailClient({
    tenantId: TENANT_ID,
    clientId: CLIENT_ID,
    clientSecret: CLIENT_SECRET,
    sharedMailbox: SHARED_MAILBOX,
    attachmentUploadOptions: {
      largeFileThreshold: 3 * 1024 * 1024,
      chunkSize: 320 * 1024
    },
    logger
  })

  // 故意触发错误：缺少收件人
  try {
    await mailClient.sendMail({ subject: 'No recipient example', to: '', text: 'x' } as any)
  } catch (e) {
    console.log('Triggered error to demonstrate custom logger')
  }
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})
