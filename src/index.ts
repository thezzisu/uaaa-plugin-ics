import { definePlugin, arktype, BusinessError } from '@uaaa/server'
import { existsSync } from 'node:fs'
import { stat, readFile } from 'node:fs/promises'
import path from 'node:path'

const tConfig = arktype.type({
  icsDir: 'string',
  icsAppId: 'string'
})

type IMyConfig = typeof tConfig.infer

declare module '@uaaa/server' {
  interface IConfig extends IMyConfig {}
  interface ICredentialTypeMap {
    iaaa: string
  }
}

class CachedObject<T> {
  private mtime = 0
  private data: T | null = null
  constructor(public filename: string, private transform?: (json: any) => T) {}

  async fetch() {
    const fstat = await stat(this.filename)
    if (!this.data || fstat.mtimeMs > this.mtime) {
      const obj = JSON.parse(await readFile(this.filename, 'utf-8'))
      this.data = this.transform ? this.transform(obj) : obj
      this.mtime = fstat.mtimeMs
    }
    return this.data!
  }
}

export default definePlugin({
  name: 'ics',
  configType: tConfig,
  setup: async (ctx) => {
    const icsDir = ctx.app.config.get('icsDir')
    const icsAppId = ctx.app.config.get('icsAppId')
    const rootDir = path.resolve(icsDir)
    const noJsonPath = path.join(rootDir, 'no.json')
    if ([rootDir, noJsonPath].some((p) => !existsSync(p))) {
      throw new Error('ICS: necessary files not found')
    }
    const noList = new CachedObject(noJsonPath, (json) => new Set<number>(json))

    ctx.app.session.hook('preDerive', async (token, targetAppId, clientAppId, securityLevel) => {
      if (clientAppId !== icsAppId) return
      const { sub } = token
      const user = await ctx.app.db.users.findOne({ _id: sub })
      const stuNo = user?.claims['iaaa:identity_id']?.value
      if (!stuNo) throw new BusinessError('FORBIDDEN', { msg: 'Contact TA to get access' })
      const noSet = await noList.fetch()
      if (!noSet.has(+stuNo)) {
        throw new BusinessError('FORBIDDEN', { msg: 'Contact TA to get access' })
      }
    })
  }
})
