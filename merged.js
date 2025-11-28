/**
 * 完全基于模板 PPT（PizZip）进行复制 / 替换 / 注入，不使用 PptxGenJS
 */

const fs = require('fs-extra')
const path = require('path')
const PizZip = require('pizzip')
const mammoth = require('mammoth')
const { fromPath: pdfFromPath } = require('pdf2pic')
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js')
const { createCanvas } = require('canvas')
const crypto = require('crypto')

pdfjsLib.GlobalWorkerOptions.workerSrc = require.resolve('pdfjs-dist/legacy/build/pdf.worker.js')

const ROOT = path.resolve(__dirname, '.')
const DATA_DIR = path.join(ROOT, 'data')
const OUTPUT_DIR = path.join(ROOT, 'output')
const PDF_IMAGE_DIR = path.join(OUTPUT_DIR, '_pdf_images')
const PDF_TEMP_DIR = path.join(PDF_IMAGE_DIR, '_tmp')

const TEMPLATE_CANDIDATES = ['2025员工体检报告（模板）.pptx', 'template.pptx']
const EMPLOYEE_SHEET_CANDIDATES = ['员工表.xlsx', 'employees.xlsx']

const IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.gif', '.bmp'])
const PDF_EXTENSIONS = new Set(['.pdf'])
const EMU_PER_INCH = 914400

const ExcelJS = require('exceljs')

async function loadEmployeesFromExcel () {
  const excelPath = await resolveExistingPath(
    EMPLOYEE_SHEET_CANDIDATES,
    '员工表'
  )

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(excelPath)

  const sheet = workbook.worksheets[0]
  if (!sheet) throw new Error('员工表读取失败：没有找到第一个 Sheet')

  // 读取表头 -> 列号映射
  const headerMap = {}
  sheet.getRow(1).eachCell((cell, col) => {
    headerMap[cell.value] = col
  })

  function col (name) {
    return headerMap[name]
  }

  const employees = []
  sheet.eachRow((row, rowIndex) => {
    if (rowIndex === 1) return

    const name = row.getCell(col('姓名')).value || ''
    if (name === null || name === undefined || name === '') return

    employees.push({
      name: String(name).trim(),
      id: String(row.getCell(col('工号')).value || '').trim(),
      gender: String(row.getCell(col('性别')).value || '').trim(),
      age: String(row.getCell(col('年龄')).value || '').trim(),
      summaryPath: null, // 会在附件扫描时自动匹配员工总结 docx
    })
  })

  return employees
}


// helper: find template
async function resolveExistingPath (candidates, label) {
  for (const candidate of candidates) {
    const fullPath = path.join(ROOT, candidate)
    if (await fs.pathExists(fullPath)) {
      if (candidate !== candidates[0]) {
        console.warn(`ℹ️ ${label} 使用备用路径：${candidate}`)
      }
      return fullPath
    }
  }
  throw new Error(`${label} 缺失，请确认以下文件之一存在：${candidates.join(', ')}`)
}

// ---------- xml/text helpers ----------
function escapeXmlValue (value) {
  if (!value) return ''
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

function escapeRegex (str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/**
 * 替换 slide XML 中的占位文本（只在 <a:t> 内替换，尽量不破坏结构）
 */
function replaceTextInSlidePreferred (xmlContent, replacements) {
  let result = xmlContent
  const optionalSpaces = '(?:\\s|\\u00A0|&nbsp;|&#160;|&#xA0;)*'

  for (const [key, value] of Object.entries(replacements)) {
    if (value === undefined || value === null) continue
    const valueStr = String(value)
    if (!valueStr && key !== '性别' && key !== '年龄') continue

    const escapedValue = escapeXmlValue(valueStr)
    const keyPattern = escapeRegex(key)

    const textRunPattern = /(<a:t[^>]*>)([\s\S]*?)(<\/a:t>)/gi
    result = result.replace(textRunPattern, (match, openTag, content, closeTag) => {
      let updatedContent = content

      const placeholderPattern = new RegExp(`\\{\\{${optionalSpaces}${keyPattern}${optionalSpaces}\\}\\}`, 'gi')
      updatedContent = updatedContent.replace(placeholderPattern, (phMatch) => {
        const suffixMatch = phMatch.match(/((?:\s|\u00A0|&nbsp;|&#160;|&#xA0;)*)$/i)
        const suffix = suffixMatch ? suffixMatch[1] : ''
        return `${escapedValue}${suffix}`
      })

      const bracketPattern = new RegExp(`\\[${optionalSpaces}${keyPattern}${optionalSpaces}\\]`, 'gi')
      updatedContent = updatedContent.replace(bracketPattern, escapedValue)

      const directPattern = new RegExp(`${optionalSpaces}${keyPattern}${optionalSpaces}`, 'gi')
      updatedContent = updatedContent.replace(directPattern, escapedValue)

      return `${openTag}${updatedContent}${closeTag}`
    })
  }

  return result
}

function parseRelationships (xml) {
  const relRegex = /<Relationship\s+([^>]+?)\/>/g
  const attrRegex = /([A-Za-z:]+)="([^"]*)"/g
  const rels = []
  let match
  while ((match = relRegex.exec(xml))) {
    const attrs = {}
    let attrMatch
    while ((attrMatch = attrRegex.exec(match[1]))) {
      attrs[attrMatch[1]] = attrMatch[2]
    }
    rels.push(attrs)
  }
  return rels
}

function buildRelationshipsXml (rels) {
  const items = rels
    .map((rel) => {
      const attrs = Object.entries(rel)
        .map(([key, value]) => `${key}="${value}"`)
        .join(' ')
      return `  <Relationship ${attrs}/>`
    })
    .join('\n')

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${items}
</Relationships>`
}

// utilities for numbering
function getNextNumber (files, regex) {
  if (!files || !files.length) {
    return 1
  }
  const numbers = files
    .map((file) => {
      const match = file.name.match(regex)
      return match ? parseInt(match[1], 10) : 0
    })
    .filter((num) => Number.isFinite(num))
  if (!numbers.length) {
    return 1
  }
  return Math.max(...numbers) + 1
}

function getInitialMediaCounter (zip) {
  const mediaFiles = zip.file(/^ppt\/media\/.+\.(png|jpg|jpeg|bmp|gif)$/i)
  if (!mediaFiles || !mediaFiles.length) {
    return 1
  }
  const numbers = mediaFiles
    .map((file) => {
      const match = file.name.match(/(\d+)\.(png|jpg|jpeg|bmp|gif)$/i)
      return match ? parseInt(match[1], 10) : 0
    })
    .filter((num) => Number.isFinite(num))
  return numbers.length ? Math.max(...numbers) + 1 : 1
}

// ensure override in [Content_Types].xml
function ensureContentTypeOnce (contentTypesText, partName, contentType) {
  const partNameRegex = new RegExp(`<Override[^>]*PartName="${escapeRegex(partName)}"[^>]*>`, 'i')
  if (contentTypesText.match(partNameRegex)) {
    return contentTypesText
  }
  const override = `  <Override PartName="${partName}" ContentType="${contentType}"/>`
  const typesEndIndex = contentTypesText.lastIndexOf('</Types>')
  if (typesEndIndex === -1) {
    throw new Error('[Content_Types].xml 缺少 </Types>')
  }
  return contentTypesText.slice(0, typesEndIndex) + `\n${override}\n` + contentTypesText.slice(typesEndIndex)
}

// ---------- asset collection (与之前逻辑兼容，简化版本) ----------
async function safeReadDir (dir) {
  if (!(await fs.pathExists(dir))) return []
  return fs.readdir(dir)
}

function normalizeName (value) {
  if (!value) return ''
  return String(value)
    .replace(/\.[^.]+$/, '')
    .replace(/[-_].*$/, '')
    .replace(/\s+/g, '')
    .toLowerCase()
}

function buildAttachmentLabel (fileName) {
  const base = path.basename(fileName, path.extname(fileName))
  const parts = base.split(/[-_]/)
  return parts.slice(1).join('-') || base
}

async function buildImageItems (assetInfo, employee) {
  const items = []
  for (const att of assetInfo.attachments) {
    if (att.type === 'image') {
      items.push({ label: att.label, fullPath: att.fullPath })
    } else if (att.type === 'pdf') {
      const converted = await convertPdfAttachment(att, employee)
      items.push(...converted)
    }
  }
  return items
}

async function extractSummaryText (filePath) {
  try {
    const { value } = await mammoth.extractRawText({ path: filePath })
    return value
      .split('\n')
      .map((i) => i.trim())
      .filter(Boolean)
      .join('\n')
  } catch (e) {
    console.warn('无法读取总结：', e.message)
    return ''
  }
}

// PDF conversions (reuse earlier logic)
async function convertPdfAttachment (pdfAttachment, employee) {
  let tempPdfPath = ''
  try {
    const safeLabel = `${employee.id}_${employee.name}_${pdfAttachment.label}`
    const tempPdfName = `${crypto.randomUUID()}.pdf`
    tempPdfPath = path.join(PDF_TEMP_DIR, tempPdfName)
    await fs.ensureDir(PDF_TEMP_DIR)
    await fs.copyFile(pdfAttachment.fullPath, tempPdfPath)

    const pdfImgPages = await convertWithPdfRenderer(tempPdfPath, pdfAttachment)
    if (pdfImgPages.length) return pdfImgPages
    return convertWithPdf2Pic(tempPdfPath, pdfAttachment, safeLabel)
  } catch (e) {
    console.warn('pdf 转图失败：', e.message)
    return []
  } finally {
    if (tempPdfPath) await fs.remove(tempPdfPath).catch(() => { })
  }
}

async function convertWithPdfRenderer (pdfPath, attachment) {
  try {
    const pdfBuffer = await fs.readFile(pdfPath)
    const pdfData = new Uint8Array(pdfBuffer)
    const pdfDoc = await pdfjsLib.getDocument({ data: pdfData, verbosity: 0 }).promise
    const pages = []
    for (let pageNumber = 1; pageNumber <= pdfDoc.numPages; pageNumber++) {
      const page = await pdfDoc.getPage(pageNumber)
      const viewport = page.getViewport({ scale: 2 })
      const canvas = createCanvas(viewport.width, viewport.height)
      const context = canvas.getContext('2d')
      await page.render({ canvasContext: context, viewport }).promise
      const imageBuffer = canvas.toBuffer('image/png')
      pages.push({
        label: `${attachment.label} 第${pageNumber}页`,
        data: `data:image/png;base64,${imageBuffer.toString('base64')}`,
      })
    }
    return pages
  } catch (e) {
    console.warn('pdf renderer 失败：', e.message)
    return []
  }
}

async function convertWithPdf2Pic (pdfPath, attachment, safeLabel) {
  try {
    await fs.ensureDir(PDF_IMAGE_DIR)
    const convert = pdfFromPath(pdfPath, {
      density: 144,
      format: 'png',
      width: 1200,
      height: 800,
      saveFilename: safeLabel,
      savePath: PDF_IMAGE_DIR,
    })
    const pages = await convert.bulk(-1, { responseType: 'base64' })
    return pages.map((p) => ({
      label: `${attachment.label} 第${p.page || '?'}页`,
      data: `data:image/png;base64,${p.base64}`,
    }))
  } catch (e) {
    console.warn('pdf2pic 失败：', e.message)
    return []
  }
}

// ---------- core: build a report per employee using template as base ----------
async function main () {
  await fs.ensureDir(OUTPUT_DIR)
  await fs.ensureDir(PDF_IMAGE_DIR)
  await fs.ensureDir(PDF_TEMP_DIR)

  const templatePath = await resolveExistingPath(TEMPLATE_CANDIDATES, '模板文件')
  const templateBuffer = await fs.readFile(templatePath)

  console.log('正在读取员工表（Excel）...')
  const employees = await loadEmployeesFromExcel()
  console.log(`员工数量：${employees.length}`)


  // scan data dir for files
  const availableFiles = await safeReadDir(DATA_DIR)

  const success = []
  for (const emp of employees) {
    try {
      // collect assets (简单匹配同名文件)
      const normalized = normalizeName(emp.name)
      const matched = availableFiles.filter((f) => normalizeName(f) === normalized)

      // summary docx?
      const summaryFile = matched.find((f) => f.includes('总结') && path.extname(f).toLowerCase() === '.docx')
      const summaryText = summaryFile ? await extractSummaryText(path.join(DATA_DIR, summaryFile)) : (emp.summary || '')

      // attachments
      const attachments = matched
        .filter((f) => f !== summaryFile)
        .map((f) => {
          const ext = path.extname(f).toLowerCase()
          const type = IMAGE_EXTENSIONS.has(ext) ? 'image' : PDF_EXTENSIONS.has(ext) ? 'pdf' : 'other'
          return {
            fileName: f,
            fullPath: path.join(DATA_DIR, f),
            label: buildAttachmentLabel(f),
            type,
            ext,
          }
        })

      const assetInfo = { attachments, summaryText }

      // build imageItems (图片或 pdf => image pages)
      const imageItems = await buildImageItems(assetInfo, emp)

      // create a fresh copy of template zip for this employee
      const templateZip = new PizZip(templateBuffer)
      const outputZip = new PizZip(templateBuffer) // start from template

      // read core files
      const presentationXml = outputZip.file('ppt/presentation.xml').asText()
      const presentationRelsXml = outputZip.file('ppt/_rels/presentation.xml.rels').asText()
      let contentTypesXml = outputZip.file('[Content_Types].xml').asText()

      // determine counters
      const slideFiles = outputZip.file(/^ppt\/slides\/slide\d+\.xml$/i) || []
      const nextSlideNumber = getNextNumber(slideFiles, /slide(\d+)\.xml/i)
      const existingSlideIds = [...presentationXml.matchAll(/<p:sldId[^>]*id="(\d+)"/g)].map(m => parseInt(m[1], 10))
      const nextSlideId = (existingSlideIds.length ? Math.max(...existingSlideIds) : 256) + 1
      const existingRelIds = [...presentationRelsXml.matchAll(/Id="rId(\d+)"/g)].map(m => parseInt(m[1], 10) || 0)
      let nextRelId = existingRelIds.length ? Math.max(...existingRelIds) + 1 : 10

      // media counter
      let mediaCounter = getInitialMediaCounter(outputZip)

      // --- 1) replace cover placeholders on slide1.xml ---
      const slide1Path = 'ppt/slides/slide1.xml'
      if (outputZip.file(slide1Path)) {
        let slide1Xml = outputZip.file(slide1Path).asText()
        const dateText = (new Date()).toISOString().slice(0, 10).replace(/-/g, '') // simple date
        const replacements = {
          姓名: emp.name || '',
          性别: emp.gender || '',
          年龄: emp.age || '',
          日期: `${new Date().getFullYear()}年${String(new Date().getMonth() + 1).padStart(2, '0')}月${String(new Date().getDate()).padStart(2, '0')}日`,
        }
        slide1Xml = replaceTextInSlidePreferred(slide1Xml, replacements)
        outputZip.file(slide1Path, slide1Xml)
      }

      // --- 2) inject summary into slide3.xml ---
      const slide3Path = 'ppt/slides/slide3.xml'
      if (outputZip.file(slide3Path) && (assetInfo.summaryText && assetInfo.summaryText.trim())) {
        let slide3Xml = outputZip.file(slide3Path).asText()
        slide3Xml = replaceTextInSlidePreferred(slide3Xml, { 总结: assetInfo.summaryText })
        outputZip.file(slide3Path, slide3Xml)
      }

      // --- 3) find template slide2 and its rels & create copies per image item ---
      const templateSlideNumber = 2
      const templateSlidePath = `ppt/slides/slide${templateSlideNumber}.xml`
      const templateSlideRelsPath = `ppt/slides/_rels/slide${templateSlideNumber}.xml.rels`
      const templateSlideFile = templateZip.file(templateSlidePath)
      const templateSlideRelsFile = templateZip.file(templateSlideRelsPath)

      if (templateSlideFile && imageItems.length) {
        const templateSlideXml = templateSlideFile.asText()
        const templateSlideRelsXml = templateSlideRelsFile ? templateSlideRelsFile.asText() : null
        // parse template slide rels to find image relationships
        const templateRels = templateSlideRelsXml ? parseRelationships(templateSlideRelsXml) : []

        // determine where to insert new slide ids in presentation.xml:
        // find the existing <p:sldId> that references slides/slide2.xml and insert after it
        const presMatch = outputZip.file('ppt/presentation.xml').asText()
        let sldIdListMatch = presMatch.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/)
        let sldIdListInner = sldIdListMatch ? sldIdListMatch[1] : ''
        const presRelsXml = outputZip.file('ppt/_rels/presentation.xml.rels').asText()
        // find rId that points to slide2
        const presRels = parseRelationships(presRelsXml)
        let slide2Rid = null
        for (const rel of presRels) {
          if (rel.Target && rel.Target.replace(/^\/?/, '') === `slides/slide${templateSlideNumber}.xml`) {
            slide2Rid = rel.Id
            break
          }
        }
        // find the index in sldIdListInner where r:id="slide2Rid"
        const sldIdLines = sldIdListInner.split('\n').map(l => l.trim()).filter(Boolean)
        let insertIndex = -1
        if (slide2Rid) {
          for (let i = 0; i < sldIdLines.length; i++) {
            if (sldIdLines[i].includes(`r:id="${slide2Rid}"`)) {
              insertIndex = i + 1
              break
            }
          }
        }
        // fallback: append at end of existing slide list if not found
        if (insertIndex === -1) {
          insertIndex = sldIdLines.length
        }

        // build new slides one by one
        const newSlideEntries = []
        const newSlideRelsToAppend = []
        let localNextSlideNumber = nextSlideNumber
        let localNextSlideId = nextSlideId
        let localNextRelId = nextRelId

        for (const image of imageItems) {
          const newSlideNumber = localNextSlideNumber++
          const newSlidePath = `ppt/slides/slide${newSlideNumber}.xml`
          const newSlideRelsPath = `ppt/slides/_rels/slide${newSlideNumber}.xml.rels`

          // 1) create a new media for this image
          // if image.data (base64) use that; else read fullPath
          let newMediaPath = null
          if (image.data) {
            const ext = (image.data.match(/^data:image\/(\w+);base64,/)?.[1] || 'png')
            const safeExt = ext.startsWith('.') ? ext : `.${ext}`
            newMediaPath = `ppt/media/image_${mediaCounter++}${safeExt}`
            const base64 = image.data.replace(/^data:image\/\w+;base64,/, '')
            const buffer = Buffer.from(base64, 'base64')
            outputZip.file(newMediaPath, buffer)
            contentTypesXml = ensureContentTypeOnce(contentTypesXml, `/${newMediaPath}`, 'image/png')
          } else if (image.fullPath) {
            const ext = path.extname(image.fullPath) || '.png'
            newMediaPath = `ppt/media/image_${mediaCounter++}${ext}`
            const buffer = await fs.readFile(image.fullPath)
            outputZip.file(newMediaPath, buffer)
            const ct = ext.match(/\.jpe?g/i) ? 'image/jpeg' : ext.match(/\.png/i) ? 'image/png' : 'image/png'
            contentTypesXml = ensureContentTypeOnce(contentTypesXml, `/${newMediaPath}`, ct)
          } else {
            continue
          }

          // 2) create new slide xml: base on templateSlideXml, replace text placeholders (影像标题, 姓名, 工号 maybe)
          const replacements = {
            影像标题: image.label || '',
            姓名: emp.name || '',
            工号: emp.id || '',
          }
          let newSlideXml = replaceTextInSlidePreferred(templateSlideXml, replacements)

          // write slide xml
          outputZip.file(newSlidePath, newSlideXml)
          // ensure content type for slide
          contentTypesXml = ensureContentTypeOnce(contentTypesXml, `/${newSlidePath}`, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')

          // 3) prepare rels: clone template relations, but for image relationships change Target to our newMediaPath
          const clonedRels = []
          if (templateRels && templateRels.length) {
            for (const r of templateRels) {
              // clone object shallow
              const nr = Object.assign({}, r)
              // if rel is image type, replace Target to newMediaPath (keeping same Id)
              if (nr.Type && nr.Type.includes('/image')) {
                // set relative path from slide to media
                // slides are in ppt/slides/, media in ppt/media/
                nr.Target = path.posix.relative(`ppt/slides`, newMediaPath).replace(/\\/g, '/')
              } else if (nr.Type && nr.Type.includes('slideLayout')) {
                // resolve layout path to template and reuse (we keep same Target)
                const absolute = resolveRelationshipPath(`ppt/slides/slide${templateSlideNumber}.xml`, nr.Target)
                // we assume layout already exists in template, map to same relative path
                nr.Target = path.posix.relative(`ppt/slides`, absolute).replace(/\\/g, '/')
              } else {
                // other types: copy as resolved relative
                const absolute = resolveRelationshipPath(`ppt/slides/slide${templateSlideNumber}.xml`, nr.Target)
                nr.Target = path.posix.relative(`ppt/slides`, absolute).replace(/\\/g, '/')
              }
              clonedRels.push(nr)
            }
          }
          // write the new rels xml for this slide
          outputZip.file(newSlideRelsPath, buildRelationshipsXml(clonedRels))

          // 4) add new presentation rel entry for this slide
          const newRelId = `rId${localNextRelId++}`
          newSlideEntries.push({
            slideNumber: newSlideNumber,
            slideId: localNextSlideId++,
            relId: newRelId,
            positionIndex: insertIndex, // we'll insert later
          })
          newSlideRelsToAppend.push(`  <Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newSlideNumber}.xml"/>`)
        }

        // 5) update presentation.xml: insert new <p:sldId> entries after the determined index
        // read current presContent
        let presContent = outputZip.file('ppt/presentation.xml').asText()
        let sldIdMatch = presContent.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/)
        const currentInner = sldIdMatch ? sldIdMatch[1].trim().split('\n').map(l => l.trim()).filter(Boolean) : []
        // build new entries
        const newEntriesXml = newSlideEntries.map((s) => `    <p:sldId id="${s.slideId}" r:id="${s.relId}"/>`)
        // insert at insertIndex
        const combined = [...currentInner.slice(0, insertIndex), ...newEntriesXml, ...currentInner.slice(insertIndex)]
        const newSldIdLst = `<p:sldIdLst>\n${combined.join('\n')}\n</p:sldIdLst>`
        presContent = presContent.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, newSldIdLst)
        outputZip.file('ppt/presentation.xml', presContent)

        // 6) update presentation.xml.rels: append new Relationship lines before </Relationships>
        let presRelsContent = outputZip.file('ppt/_rels/presentation.xml.rels').asText()
        const insertPos = presRelsContent.lastIndexOf('</Relationships>')
        presRelsContent = presRelsContent.slice(0, insertPos) + '\n' + newSlideRelsToAppend.join('\n') + '\n' + presRelsContent.slice(insertPos)
        outputZip.file('ppt/_rels/presentation.xml.rels', presRelsContent)

        // 7) update [Content_Types].xml was handled per-media & per-slide above in contentTypesXml
        outputZip.file('[Content_Types].xml', contentTypesXml)

      } // end processing template slide2 copies

      // write the final zip buffer to a new pptx file for this employee
      const outFileName = `员工体检报告_${(emp.name || 'unknown')}_${(emp.id || '')}_${(new Date()).toISOString().replace(/[:.]/g, '')}.pptx`
      const outPath = path.join(OUTPUT_DIR, outFileName)

      const finalBuffer = outputZip.generate({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } })
      await fs.writeFile(outPath, finalBuffer)

      success.push({ emp, outPath })
      console.log(`✓ 已生成：${outPath}`)
    } catch (err) {
      console.error(`❌ ${emp.name || 'UNKNOWN'} 生成失败：`, err)
    }
  }

  console.log('生成完成，数量：', success.length)
}

function resolveRelationshipPath (from, target) {
  // from e.g. "ppt/slides/slide2.xml", target e.g. "../media/image1.png"
  const baseDir = path.posix.dirname(from)
  return path.posix.normalize(path.posix.join(baseDir, target))
}

// ================= run
main().catch((e) => {
  console.error('执行失败：', e)
  process.exit(1)
})
