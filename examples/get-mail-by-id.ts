import { MailClient } from '../src'

const TENANT_ID = process.env.TENANT_ID || ''
const CLIENT_ID = process.env.CLIENT_ID || ''
const CLIENT_SECRET = process.env.CLIENT_SECRET || ''
const SHARED_MAILBOX = process.env.SHARED_MAILBOX || ''

const id = process.env.MESSAGE_ID || ''

async function main() {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !SHARED_MAILBOX) {
    throw new Error('Please set TENANT_ID, CLIENT_ID, CLIENT_SECRET, SHARED_MAILBOX environment variables')
  }

  const mailClient = new MailClient({
    tenantId: TENANT_ID,
    clientId: CLIENT_ID,
    clientSecret: CLIENT_SECRET,
    sharedMailbox: SHARED_MAILBOX,
    attachmentUploadOptions: {
      largeFileThreshold: 3 * 1024 * 1024,
      chunkSize: 320 * 1024
    }
  })

  if (!id) throw new Error('Please set MESSAGE_ID to the target email id')

  const msg = await mailClient.getMailById(id, { includeAttachments: true })
  console.log('Message:', {
    id: msg.id,
    subject: msg.subject,
    attachments: msg.attachments?.length || 0
  })
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})
