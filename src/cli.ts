#!/usr/bin/env node
import { Builtins, Cli, Command, Option } from 'clipanion'
import { readFile, writeFile } from 'node:fs/promises'
import { oraPromise } from 'ora'
import * as XLSX from 'xlsx'

class GenerateNoListCommand extends Command {
  static paths = [[`generate-no-list`]]
  fmbd = Option.String(`-f,--fmbd`, { required: true, description: 'Fen Ban Ming Dan file path' })
  additional = Option.String(`-a,--add`, { required: false, description: 'Additional file path' })
  output = Option.String(`-o,--output`, { required: true, description: 'Output file path' })

  async execute() {
    const studentNo = new Set<number>()
    await oraPromise(
      async (spinner) => {
        const binary = await readFile(this.fmbd)
        const workbook = XLSX.read(binary, { type: 'buffer' })

        function parseCell(cell: XLSX.CellObject) {
          if (cell?.t !== 'n') return null
          const num = cell.v
          if (typeof num !== 'number') return null
          if (!Number.isSafeInteger(num)) return null
          if (num <= 0) return null
          return num
        }

        for (const sheetName of workbook.SheetNames) {
          spinner.text = `parsing sheet ${sheetName}`
          const workSheet = workbook.Sheets[sheetName]
          for (let row = 4; ; row++) {
            const cell = workSheet[`A${row}`]
            const id = parseCell(cell)
            if (id === null) break
            studentNo.add(id)
          }
        }
      },
      { prefixText: 'Loading Fen Ban Ming Dan:', text: 'parsing file' }
    )

    await oraPromise(
      async (spinner) => {
        if (!this.additional) return
        const file = await readFile(this.additional, 'utf-8')
        // Format: Section Name No
        for (const line of file.split('\n')) {
          const no = line.split(/\s+/)[2]
          const id = parseInt(no)
          if (Number.isSafeInteger(id) && id > 0) {
            studentNo.add(id)
          }
        }
      },
      { prefixText: 'Loading Additional File:', text: 'parsing file' }
    )

    console.log(`Loaded ${studentNo.size} student numbers`)
    const json = JSON.stringify(
      [...studentNo].sort((a, b) => a - b),
      null,
      2
    )
    await oraPromise(
      async () => {
        await writeFile(this.output, json)
      },
      { prefixText: 'Writing to file:', text: this.output }
    )
  }
}

class ParseTACommand extends Command {
  static paths = [[`parse-ta`]]
  fmbd = Option.String(`-f,--fmbd`, { required: true, description: 'Fen Ban Ming Dan file path' })
  courseId = Option.String(`-c,--course-id`, { required: true, description: 'Course ID' })
  cookie = Option.String(`-k,--cookie`, { required: true, description: 'Cookie' })

  async execute() {
    const binary = await readFile(this.fmbd)
    const workbook = XLSX.read(binary, { type: 'buffer' })

    function parseCell(cell: XLSX.CellObject) {
      if (cell?.t !== 'n') return null
      const num = cell.v
      if (typeof num !== 'number') return null
      if (!Number.isSafeInteger(num)) return null
      if (num <= 0) return null
      return num
    }
    let output = ''
    for (const sheetName of workbook.SheetNames) {
      const workSheet = workbook.Sheets[sheetName]
      const cell = workSheet[`D1`]
      const sectionNo = sheetName.replace(/[^0-9]/g, '')
      const taName = cell?.v
      console.log(`Processing ${sectionNo} ${taName}`)
      const resp = await fetch(
        `https://course.pku.edu.cn/webapps/blackboard/execute/userManager?course_id=${this.courseId}`,
        {
          headers: { 'content-type': 'application/x-www-form-urlencoded', cookie: this.cookie },
          body: new URLSearchParams({
            actionString: 'Search',
            navType: 'cp_list_modify_users',
            enroll: '',
            objectVar: 'tmpuser',
            searchTypeString: 'UserInformation',
            userInfoSearchKeyString: 'GivenName',
            userInfoSearchOperatorString: 'Equals',
            userInfoSearchText: taName
          }),
          method: 'POST'
        }
      )
      const html = await resp.text()
      const match = html.match(
        /<span class="profileCardAvatarThumb">[\s\S]*?(\d{10})[\s\S]*?<\/span>/
      )
      if (!match) {
        console.error(`TA ${taName} not found in course ${this.courseId}`)
        continue
      }
      const taId = match[1]
      output += `${sectionNo} ${taName} ${taId}\n`
    }
    console.log()
    console.log(output)
  }
}

const cli = new Cli({
  binaryName: `uaaa-plugin-ics`,
  binaryLabel: `ICS plugin for UAAA CLI`
})
cli.register(Builtins.HelpCommand)
cli.register(GenerateNoListCommand)
cli.register(ParseTACommand)
cli.runExit(process.argv.slice(2))
