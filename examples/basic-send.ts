import { MailClient } from '../src'

const TENANT_ID = process.env.TENANT_ID || ''
const CLIENT_ID = process.env.CLIENT_ID || ''
const CLIENT_SECRET = process.env.CLIENT_SECRET || ''
const SHARED_MAILBOX = process.env.SHARED_MAILBOX || ''

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

  const mailId = await mailClient.sendMail({
    subject: 'Hello from msgraph-shared-mailbox',
    to: process.env.TO || 'someone@example.com',
    text: 'This is a text body. You can also use HTML.'
  })
  console.log('mail message id:', mailId)
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})
