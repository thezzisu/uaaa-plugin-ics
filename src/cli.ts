#!/usr/bin/env node
import { Builtins, Cli, Command, Option } from 'clipanion'
import { readFile, writeFile } from 'node:fs/promises'
import { oraPromise } from 'ora'
import * as XLSX from 'xlsx'
import ky from 'ky'
import jsdom from 'jsdom'

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

class SyncAutolabCommand extends Command {
  static paths = [[`sync-autolab`]]
  courseName = Option.String(`-c,--course-name`, { required: true, description: 'Course Name' })
  cookie = Option.String(`-k,--cookie`, { required: true, description: 'Cookie' })
  fmbd = Option.String(`-f,--fmbd`, { required: true, description: 'Fen Ban Ming Dan file path' })
  additional = Option.String(`-a,--add`, { required: true, description: 'Additional file path' })

  async execute() {
    const sectionMap = new Map<number, string>()
    const courseAssistant = new Set<number>()
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
      const workSheet = workbook.Sheets[sheetName]
      const sectionNo = +sheetName.replace(/[^0-9]/g, '')
      // console.log(`Processing section ${sectionNo}`)
      for (let row = 4; ; row++) {
        const cell = workSheet[`A${row}`]
        const id = parseCell(cell)
        if (id === null) break
        sectionMap.set(id, sectionNo.toString())
      }
    }

    const file = await readFile(this.additional, 'utf-8')
    for (const line of file.split('\n')) {
      const [section, name, no, flag] = line.split(/\s+/)
      const id = parseInt(no)
      if (Number.isSafeInteger(id) && id > 0) {
        sectionMap.set(id, section)
        if (flag === '1') {
          courseAssistant.add(id)
        }
      }
    }

    console.log(`Loaded ${sectionMap.size} section mappings`)

    const http = ky.extend({
      prefixUrl: `https://autolab.pku.edu.cn/courses/${this.courseName}`,
      headers: { cookie: this.cookie }
    })
    const usersDom = new jsdom.JSDOM(await http.get('users').text())
    const elements = [
      ...usersDom.window.document.querySelectorAll('#ajaxTable > table > tbody > tr.user-row')
    ]
    console.log(`Loaded ${elements.length} users`)
    for (const tr of elements) {
      const a = tr.querySelector('td:nth-child(4)>a') as HTMLAnchorElement
      if (a) {
        const email = a.textContent ?? ''
        const no = +email.split('@')[0]
        const id = +a.href.split('course_user_data/')[1]
        if (email && no && id) {
          const section = tr.querySelector('td:nth-child(7)')?.textContent ?? ''
          if (!section) {
            console.log(`Processing email=${email} no=${no} id=${id}`)
            const sectionNo = sectionMap.get(no)
            if (!sectionNo) {
              console.warn(`No section found for ${no}`)
              continue
            }
            console.log(`Setting section to ${sectionNo}`)
            const html = await http.get(`course_user_data/${id}/edit`).text()
            const dom = new jsdom.JSDOM(html)
            const tokenInput = dom.window.document.querySelector(
              'input[name="authenticity_token"]'
            ) as HTMLInputElement
            const userIdInput = dom.window.document.querySelector(
              'input[name="course_user_datum[user_attributes][id]"]'
            ) as HTMLInputElement
            const nicknameInput = dom.window.document.querySelector(
              'input[name="course_user_datum[nickname]"]'
            ) as HTMLInputElement
            const resp = await http
              .post(`course_user_data/${id}`, {
                body: new URLSearchParams({
                  utf8: '✓',
                  _method: 'patch',
                  authenticity_token: tokenInput.value,
                  'course_user_datum[user_attributes][id]': userIdInput.value,
                  'course_user_datum[nickname]': nicknameInput.value,
                  'course_user_datum[course_number]': '',
                  'course_user_datum[lecture]': '',
                  'course_user_datum[section]': sectionNo,
                  'course_user_datum[tweak_attributes][value]': '0',
                  'course_user_datum[tweak_attributes][kind]': 'points',
                  'course_user_datum[dropped]': '0',
                  'course_user_datum[instructor]': '0',
                  'course_user_datum[course_assistant]': courseAssistant.has(no) ? '1' : '0'
                })
              })
              .text()
          }
        }
      }
    }
  }
}

const cli = new Cli({
  binaryName: `uaaa-plugin-ics`,
  binaryLabel: `ICS plugin for UAAA CLI`
})
cli.register(Builtins.HelpCommand)
cli.register(GenerateNoListCommand)
cli.register(ParseTACommand)
cli.register(SyncAutolabCommand)
cli.runExit(process.argv.slice(2))
