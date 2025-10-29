import { MailClient } from '../src/MailClient'
import * as fsAsync from 'node:fs/promises'
import type { CredentialOptions } from '../src'

const clientCalls: any[] = []
const responses: Record<string, any> = {}

const fakeClient = {
  calls: clientCalls,
  api: (path: string) => {
    const call: any = { path }
    const chain: any = {
      filter: (q: string) => {
        call.filter = q
        call.path = `${path}?$filter=${encodeURIComponent(q)}`
        return chain
      },
      select: (s: string) => {
        call.select = s
        return chain
      },
      get: async () => {
        call.method = 'GET'
        clientCalls.push(call)
        if (responses[call.path] !== undefined) return responses[call.path]
        if (responses[path] !== undefined) return responses[path]
        return {}
      },
      post: async (body: any) => {
        call.method = 'POST'
        call.body = body
        clientCalls.push(call)
        if (path.includes('/messages/') && path.endsWith('/send')) {
          return {}
        }
        if (path.includes('/messages/') && path.endsWith('/attachments/createUploadSession')) {
          return { uploadUrl: 'https://upload.example/session' }
        }
        if (path.includes('/messages') && !path.endsWith('/attachments') && !path.endsWith('/attachments/createUploadSession')) {
          return { id: 'MSG-1', internetMessageId: '<IM-MSG-1@exch.example.com>' }
        }
        return {}
      }
    }
    return chain
  }
}

jest.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    initWithMiddleware: jest.fn(() => fakeClient)
  }
}))

jest.mock('@azure/identity', () => ({
  ClientSecretCredential: jest.fn().mockImplementation(() => ({
    getToken: jest.fn().mockResolvedValue({ token: 'fake-token' })
  }))
}))

jest.mock('node:fs/promises', () => ({
  stat: jest.fn(async () => ({ isFile: () => true, size: 1024 })),
  readFile: jest.fn(async (_p: string) => Buffer.from('file')),
  open: jest.fn(async () => ({
    createReadStream: (_opts?: any) => ({
      async *[Symbol.asyncIterator]() {},
      close: () => {}
    }),
    close: jest.fn()
  }))
}))

const fetchMock = jest.fn(async (url: string, init?: any) => {
  if (init?.method === 'HEAD') {
    return {
      ok: true,
      headers: { get: (k: string) => (k.toLowerCase() === 'content-length' ? '10' : null) }
    } as any
  }
  if (init?.method === 'PUT') {
    return { ok: true, status: 202 } as any
  }
  return {
    ok: true,
    headers: { get: (_k: string) => null },
    arrayBuffer: async () => Buffer.from('abc'),
    body: undefined
  } as any
})
;(global as any).fetch = fetchMock

describe('MailClient', () => {
  const sharedMailbox = 'shared@example.com'
  const options: CredentialOptions = {
    sharedMailbox,
    tenantId: 'tenant',
    clientId: 'client',
    clientSecret: 'secret',
    attachmentUploadOptions: {
      largeFileThreshold: 1024,
      chunkSize: 320 * 1024
    }
  }

  beforeEach(() => {
    clientCalls.length = 0
    Object.keys(responses).forEach((k) => delete responses[k])
    fetchMock.mockClear()
  })

  test('sendMail builds recipients and text body, sends draft', async () => {
    const mc = new MailClient(options)
    const idOrObj = await mc.sendMail({
      subject: 'Subject',
      to: 'a@ex.com; b@ex.com, a@ex.com',
      text: 'Hello'
    })

    const id = typeof idOrObj === 'string' ? idOrObj : idOrObj.id
    expect(id).toBe('MSG-1')
    if (typeof idOrObj !== 'string') {
      expect(idOrObj.internetMessageId).toBeDefined()
    }

    const draftCall = clientCalls.find((c) => c.path === `/users/${sharedMailbox}/messages` && c.method === 'POST')
    expect(draftCall).toBeTruthy()
    expect(draftCall.body.subject).toBe('Subject')
    expect(draftCall.body.toRecipients.map((r: any) => r.emailAddress.address)).toEqual(['a@ex.com', 'b@ex.com'])
    expect(draftCall.body.body.contentType).toBe('text')

    const sendCall = clientCalls.find((c) => c.path === `/users/${sharedMailbox}/messages/MSG-1/send` && c.method === 'POST')
    expect(sendCall).toBeTruthy()
  })

  test('sendMail supports html body', async () => {
    const mc = new MailClient(options)
    await mc.sendMail({
      subject: 'Hi',
      to: 'x@ex.com',
      html: '<b>hi</b>'
    })
    const draftCall = clientCalls.find((c) => c.path === `/users/${sharedMailbox}/messages` && c.method === 'POST')
    expect(draftCall.body.body.contentType).toBe('html')
    expect(draftCall.body.body.content).toBe('<b>hi</b>')
  })

  test('sendMail throws when no "to" recipients', async () => {
    const mc = new MailClient(options)
    await expect(mc.sendMail({ subject: 'S', to: '', text: 'x' } as any)).rejects.toThrow('No recipients provided in "to"')
  })

  test('uses custom logger when error occurs', async () => {
    const logger = { error: jest.fn() }
    const mc = new MailClient({ ...options, logger })
    await expect(mc.sendMail({ subject: 'S', to: '', text: 'x' } as any)).rejects.toThrow('No recipients provided in "to"')
    expect(logger.error).toHaveBeenCalledTimes(1)
    const args = (logger.error as any).mock.calls[0]
    expect(args[0]).toBe('Failed to send mail')
    expect(args[1]).toBeInstanceOf(Error)
  })

  test('uploads small buffer attachment via attachments endpoint', async () => {
    const mc = new MailClient(options)
    await mc.sendMail({
      subject: 'Attach',
      to: 'x@ex.com',
      text: 'x',
      attachments: [
        {
          filename: 'note.txt',
          content: Buffer.from('hello')
        }
      ]
    })
    const attCall = clientCalls.find((c) => c.path === `/users/${sharedMailbox}/messages/MSG-1/attachments` && c.method === 'POST')
    expect(attCall).toBeTruthy()
    expect(attCall.body['@odata.type']).toBe('#microsoft.graph.fileAttachment')
    expect(attCall.body.name).toBe('note.txt')
    expect(attCall.body.contentType).toBe('text/plain')
    expect(typeof attCall.body.contentBytes).toBe('string')
  })

  test('getMailById returns message and merges attachments when requested', async () => {
    const mc = new MailClient(options)
    const messageId = 'M-123'
    responses[`/users/${sharedMailbox}/messages/${messageId}`] = { id: messageId, subject: 'S' }
    responses[`/users/${sharedMailbox}/messages/${messageId}/attachments`] = { value: [{ id: 'A1', name: 'f.txt' }] }

    const msg = await mc.getMailById(messageId, { includeAttachments: true })
    expect(msg.id).toBe(messageId)
    expect(msg.attachments).toEqual([{ id: 'A1', name: 'f.txt' }])

    const getCall = clientCalls.find((c) => c.path === `/users/${sharedMailbox}/messages/${messageId}` && c.method === 'GET')
    expect(getCall.select).toContain('subject')
    const attGetCall = clientCalls.find((c) => c.path === `/users/${sharedMailbox}/messages/${messageId}/attachments` && c.method === 'GET')
    expect(attGetCall).toBeTruthy()
  })

  test('getMailById respects custom select', async () => {
    const mc = new MailClient(options)
    const messageId = 'M-456'
    responses[`/users/${sharedMailbox}/messages/${messageId}`] = { id: messageId, subject: 'S' }
    const msg = await mc.getMailById(messageId, { select: ['id', 'subject'] })
    expect(msg.id).toBe(messageId)
    const call = clientCalls.find((c) => c.path === `/users/${sharedMailbox}/messages/${messageId}` && c.method === 'GET')
    expect(call.select).toBe('id,subject')
  })

  test('getMailById supports internetMessageId with attachments', async () => {
    const mc = new MailClient(options)
    const internetId = '<IM-123@exch.example.com>'
    const filter = `internetMessageId eq '${internetId}'`
    const listPath = `/users/${sharedMailbox}/messages?$filter=${encodeURIComponent(filter)}`
    responses[listPath] = { value: [{ id: 'M-789', subject: 'S', internetMessageId: internetId }] }
    responses[`/users/${sharedMailbox}/messages/M-789/attachments`] = { value: [{ id: 'A1', name: 'f.txt' }] }

    const msg = await mc.getMailById(internetId, { includeAttachments: true })
    expect(msg.id).toBe('M-789')
    expect(msg.internetMessageId).toBe(internetId)
    expect(msg.attachments).toEqual([{ id: 'A1', name: 'f.txt' }])

    const listCall = clientCalls.find((c) => c.path === listPath && c.method === 'GET')
    expect(listCall).toBeTruthy()
  })

  test('rejects path attachment larger than 150MB', async () => {
    const mc = new MailClient(options)
    ;(fsAsync.stat as unknown as jest.Mock).mockResolvedValue({
      isFile: () => true,
      size: 150 * 1024 * 1024 + 1
    })
    await expect(
      mc.sendMail({
        subject: 'BigPath',
        to: 'x@ex.com',
        text: 'x',
        attachments: [{ filename: 'big.bin', path: '/tmp/big.bin' }]
      })
    ).rejects.toThrow(/150MB/)
    ;(fsAsync.stat as unknown as jest.Mock).mockResolvedValue({ isFile: () => true, size: 1024 })
  })

  test('rejects href attachment larger than 150MB via HEAD', async () => {
    const mc = new MailClient(options)
    fetchMock.mockImplementationOnce(async (_url: string, init?: any) => {
      if (init?.method === 'HEAD') {
        return {
          ok: true,
          headers: { get: (k: string) => (k.toLowerCase() === 'content-length' ? String(150 * 1024 * 1024 + 1) : null) }
        } as any
      }
      return { ok: true, headers: { get: () => null }, arrayBuffer: async () => Buffer.from('x') } as any
    })

    await expect(
      mc.sendMail({
        subject: 'BigHref',
        to: 'x@ex.com',
        text: 'x',
        attachments: [{ filename: 'big.bin', href: 'http://example.com/big.bin' }]
      })
    ).rejects.toThrow(/150MB/)
  })
})
