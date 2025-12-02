const path = require('path')
const fs = require('fs-extra')
const ExcelJS = require('exceljs')
const PizZip = require('pizzip')
const PptxGenJS = require('pptxgenjs')
const mammoth = require('mammoth')
const { fromPath: pdfFromPath } = require('pdf2pic')
const { createCanvas } = require('canvas')
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js')
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

const DEFAULT_LAYOUT = {
  width: 10,
  height: 5.625,
  orientation: 'landscape',
}

function normalizeName (value, id = '') {
  if (!value) return ''
  const namePart = String(value)
    .replace(/\.[^.]+$/, '')
    .replace(/\s+/g, '')
    .toLowerCase()
  if (id) {
    return namePart + id.toLowerCase().replace(/\s+/g, '')
  }
  const match = namePart.match(/^([^-]+)-([^-]+)/)
  if (match) {
    return match[1] + match[2]
  }
  return namePart.replace(/[-_].*$/, '')
}

function normalizeText (value) {
  if (value === undefined || value === null) return ''
  return String(value).trim()
}

function sanitizeForFilename (value) {
  return value.replace(/[<>:"/\\|?*]/g, '').replace(/\s+/g, '')
}

function buildAsciiSafeLabel (value) {
  const sanitized = sanitizeForFilename(value).replace(/[^\x00-\x7F]/g, '')
  if (sanitized) {
    return sanitized
  }
  return `pdf_${Date.now()}`
}

function formatDate (date) {
  const hh = String(date.getHours()).padStart(2, '0')
  const mi = String(date.getMinutes()).padStart(2, '0')
  return `${hh}${mi}`
}

async function resolveExistingPath (candidates, label) {
  for (const candidate of candidates) {
    const fullPath = path.join(ROOT, candidate)
    if (await fs.pathExists(fullPath)) {
      if (candidate !== candidates[0]) {
        console.warn(`ℹ️ ${label}使用备用路径：${candidate}`)
      }
      return fullPath
    }
  }
  throw new Error(`${label}缺失，请确认以下文件之一存在：${candidates.join(', ')}`)
}

async function loadTemplateLayout (templatePath) {
  const fallback = { ...DEFAULT_LAYOUT }
  if (!(await fs.pathExists(templatePath))) {
    return fallback
  }
  try {
    const buffer = await fs.readFile(templatePath)
    const zip = new PizZip(buffer)
    const presentationFile = zip.file('ppt/presentation.xml')
    if (!presentationFile) {
      return fallback
    }
    const xml = presentationFile.asText()
    const sizeMatch = xml.match(/<p:sldSz[^>]*cx=["'](\d+)["'][^>]*cy=["'](\d+)["']/)
    if (!sizeMatch) {
      return fallback
    }
    const widthEmu = parseInt(sizeMatch[1], 10)
    const heightEmu = parseInt(sizeMatch[2], 10)
    if (Number.isNaN(widthEmu) || Number.isNaN(heightEmu)) {
      return fallback
    }
    const width = Number((widthEmu / EMU_PER_INCH).toFixed(3))
    const height = Number((heightEmu / EMU_PER_INCH).toFixed(3))
    const orientation = width >= height ? 'landscape' : 'portrait'
    return { width, height, orientation }
  } catch (error) {
    console.warn(`⚠️ 模板布局解析失败，将使用默认 16:9 布局。原因：${error.message}`)
    return fallback
  }
}

async function loadEmployees (sheetPath) {
  if (!(await fs.pathExists(sheetPath))) {
    throw new Error(`未找到员工表：${sheetPath}`)
  }
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(sheetPath)
  const worksheet = workbook.worksheets[0]
  if (!worksheet) {
    return []
  }
  const headerRow = worksheet.getRow(1)
  const headers = headerRow.values.map((value) => (typeof value === 'string' ? value.trim() : value))
  const employees = []
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return
    const record = {}
    row.values.forEach((cell, cellIndex) => {
      const header = headers[cellIndex]
      if (!header || typeof header !== 'string') return
      record[header] = typeof cell === 'string' ? cell.trim() : cell
    })
    const name = normalizeText(record['姓名'] || record['name'])
    if (!name) return
    employees.push({
      name,
      gender: normalizeText(record['性别']) || '未知',
      age: normalizeText(record['年龄']) || '未知',
      id: normalizeText(record['证件号']) || '未知',
      date: normalizeText(record['体检日期']) || '未知',
      raw: record,
    })
  })
  return employees
}

async function safeReadDir (dir) {
  if (!(await fs.pathExists(dir))) {
    console.warn(`⚠️ 数据目录不存在：${dir}`)
    return []
  }
  return fs.readdir(dir)
}

async function collectEmployeeAssets (employee, files) {
  const normalized = normalizeName(employee.name, employee.id)
  const matchingFiles = files.filter((file) => normalizeName(file) === normalized)
  const summaryFile = matchingFiles.find(
    (file) => file.includes('总结') && path.extname(file).toLowerCase() === '.docx',
  )
  const summaryText = summaryFile
    ? await extractSummaryText(path.join(DATA_DIR, summaryFile))
    : ''
  const attachments = matchingFiles
    .filter((file) => file !== summaryFile)
    .map((file) => {
      const ext = path.extname(file).toLowerCase()
      const type = IMAGE_EXTENSIONS.has(ext) ? 'image' : PDF_EXTENSIONS.has(ext) ? 'pdf' : 'other'
      const label = buildAttachmentLabel(file)
      const lower = file.toLowerCase()
      let priority = 99
      if (lower.includes('inbody')) {
        priority = 1
      } else if (lower.includes('尿常规') || lower.includes('尿检')) {
        priority = 2
      } else if (lower.includes('血常规')) {
        priority = 3
      } else if (lower.includes('生化检查')) {
        priority = 4
      } else if (lower.includes('糖化血红蛋白')) {
        priority = 5
      } else if (lower.includes('心电图')) {
        priority = 6
      } else if (lower.includes('ai解读')) {
        priority = 7
      }
      let category = 'other'
      if (lower.includes('inbody')) {
        category = 'inbody'
      } else if (lower.includes('心电图')) {
        category = 'ecg'
      } else if (lower.includes('尿常规') || lower.includes('尿检') || lower.includes('血常规') || lower.includes('生化检查') || lower.includes('糖化血红蛋白')) {
        category = 'lab'
      } else if (lower.includes('ai解读')) {
        category = 'ai'
      }
      return {
        fileName: file,
        fullPath: path.join(DATA_DIR, file),
        label,
        type,
        ext,
        priority,
        category,
      }
    })
    .sort((a, b) => a.priority - b.priority)
  return {
    employee,
    summaryFile: summaryFile ? path.join(DATA_DIR, summaryFile) : null,
    summaryText: summaryText || '',
    attachments,
    totalFiles: matchingFiles.length,
  }
}

async function extractSummaryText (filePath) {
  try {
    const { value } = await mammoth.extractRawText({ path: filePath })
    return value
      .split('\n')
      .map((item) => item.trim())
      .filter(Boolean)
      .join('\n')
  } catch (error) {
    console.warn(`⚠️ 无法读取AI总结：${filePath}，原因：${error.message}`)
    return ''
  }
}

function buildAttachmentLabel (fileName) {
  const base = path.basename(fileName, path.extname(fileName))
  const parts = base.split(/[-_]/)
  return parts.slice(1).join('-') || '附件'
}

async function buildImageItems (assetInfo, employee) {
  const imageItems = []
  for (const attachment of assetInfo.attachments) {
    if (attachment.type === 'image') {
      imageItems.push({
        label: attachment.label,
        fullPath: attachment.fullPath,
        category: attachment.category,
      })
    } else if (attachment.type === 'pdf') {
      const converted = await convertPdfAttachment(attachment, employee)
      imageItems.push(...converted)
    }
  }
  return imageItems
}

async function convertPdfAttachment (pdfAttachment, employee) {
  let tempPdfPath = ''
  try {
    const safeLabel = buildAsciiSafeLabel(`${employee.id}_${employee.name}_${pdfAttachment.label}`)
    const tempPdfName = `${crypto.randomUUID()}.pdf`
    tempPdfPath = path.join(PDF_TEMP_DIR, tempPdfName)
    await fs.ensureDir(PDF_TEMP_DIR)
    await fs.copyFile(pdfAttachment.fullPath, tempPdfPath)
    const pagesFromPic = await convertWithPdf2Pic(tempPdfPath, pdfAttachment, safeLabel)
    if (pagesFromPic.length) {
      return pagesFromPic.map((p) => ({ ...p, category: pdfAttachment.category }))
    }
    console.warn(`⚠️ pdf2pic 转换未返回结果，尝试使用 pdfjs-dist 兜底...`)
    const pdfImgPages = await convertWithPdfRenderer(tempPdfPath, pdfAttachment)
    if (pdfImgPages.length) {
      return pdfImgPages.map((p) => ({ ...p, category: pdfAttachment.category }))
    }
    return []
  } catch (error) {
    console.warn(`⚠️ PDF 转图失败（${pdfAttachment.fileName}）：${error.message}`)
    return []
  } finally {
    if (tempPdfPath) {
      await fs.remove(tempPdfPath).catch(() => { })
    }
  }
}

async function convertWithPdfRenderer (pdfPath, attachment) {
  try {
    const pdfBuffer = await fs.readFile(pdfPath)
    const pdfData = new Uint8Array(pdfBuffer)
    const pdfDoc = await pdfjsLib.getDocument({
      data: pdfData,
      verbosity: 0,
      cMapUrl: './node_modules/pdfjs-dist/cmaps/',
      cMapPacked: true,
    }).promise
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
  } catch (error) {
    console.warn(`⚠️ 内置 PDF 渲染失败（${attachment.fileName}）：${error.message}`)
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
    return pages.map((pageResult, idx) => ({
      label: `${attachment.label} 第${pageResult.page || idx + 1}页`,
      data: `data:image/png;base64,${pageResult.base64}`,
    }))
  } catch (error) {
    console.warn(`⚠️ pdf2pic 转换失败（${attachment.fileName}）：${error.message}`)
    return []
  }
}

function initializePresentation (layout) {
  const pptx = new PptxGenJS()
  const layoutName = `TEMPLATE_${layout.width}x${layout.height}`
  pptx.defineLayout({
    name: layoutName,
    width: layout.width,
    height: layout.height,
  })
  pptx.layout = layoutName
  pptx.author = 'Health Manage AI'
  pptx.company = 'Health Manage AI'
  pptx.subject = '员工体检报告'
  pptx.title = '2025 员工体检报告'
  return pptx
}

function buildReportFileName (employee, suffix = '') {
  const safeName = sanitizeForFilename(employee.name)
  const dateStr = formatDate(new Date())
  if (suffix) {
    // return `员工体检报告_${safeName}${suffix}_${dateStr}.pptx`
    return `体检报告_${safeName}${suffix}.pptx`
  }
  return `体检报告_${safeName}.pptx`
}

const CONTENT_TYPES = {
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  slideLayout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  slideMaster: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
}

const REL_TYPES = {
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
}

async function insertTemplateSlides (templatePath, outputPath, options = {}) {
  const [templateBuffer, outputBuffer] = await Promise.all([
    fs.readFile(templatePath),
    fs.readFile(outputPath),
  ])
  const templateZip = new PizZip(templateBuffer)
  const outputZip = new PizZip(outputBuffer)
  const slidesToCopy = [
    { templateSlide: 1, position: 'start' },
    { templateSlide: 6, position: 'end' },
    { templateSlide: 7, position: 'end' },
  ]
  const state = initializeState(outputZip, options)
  for (const config of slidesToCopy) {
    copyTemplateSlide({
      templateZip,
      outputZip,
      templateSlideNumber: config.templateSlide,
      position: config.position,
      state,
      options,
    })
  }
  if (!state.newSlides.length) {
    return
  }
  updatePresentationDocuments(outputZip, state)
  const updatedBuffer = outputZip.generate({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
  await fs.writeFile(outputPath, updatedBuffer)
}

function initializeState (zip, options) {
  const presentationXml = zip.file('ppt/presentation.xml')
  const presentationRels = zip.file('ppt/_rels/presentation.xml.rels')
  const contentTypes = zip.file('[Content_Types].xml')
  if (!presentationXml || !presentationRels || !contentTypes) {
    throw new Error('目标 PPT 结构不完整，缺少 presentation 元数据')
  }
  const presContent = presentationXml.asText()
  const presRelsContent = presentationRels.asText()

  // 修复：使用更宽松的正则匹配 ID，支持单双引号
  const slideIdMatches = [...presContent.matchAll(/<p:sldId[^>]*id=["'](\d+)["']/g)]
  const existingSlideIds = new Set(
    slideIdMatches.map((m) => parseInt(m[1], 10))
  )

  const relIdMatches = [...presRelsContent.matchAll(/Id=["']rId(\d+)["']/g)]
  const existingRelIds = new Set(
    relIdMatches.map((m) => parseInt(m[1], 10))
  )

  const masterIdMatches = [...presContent.matchAll(/<p:sldMasterId[^>]*id=["'](\d+)["']/g)]
  const existingMasterIds = new Set(
    masterIdMatches.map((m) => parseInt(m[1], 10))
  )

  const slideFiles = zip.file(/^ppt\/slides\/slide\d+\.xml$/i)
  const layoutFiles = zip.file(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/i)
  const masterFiles = zip.file(/^ppt\/slideMasters\/slideMaster\d+\.xml$/i)
  const themeFiles = zip.file(/^ppt\/theme\/theme\d+\.xml$/i)

  let nextSlideId = 256
  while (existingSlideIds.has(nextSlideId)) {
    nextSlideId++
  }
  let nextRelId = 1
  while (existingRelIds.has(nextRelId)) {
    nextRelId++
  }
  let nextMasterId = 2147483648 // Master IDs usually start high
  if (existingMasterIds.size > 0) {
    nextMasterId = Math.max(...existingMasterIds) + 1
  }

  return {
    presentationXml,
    presContent,
    presentationRels,
    presRelsContent,
    contentTypes,
    contentTypesContent: contentTypes.asText(),
    nextSlideNumber: getNextNumber(slideFiles, /slide(\d+)\.xml/i),
    nextSlideId,
    nextRelId,
    nextLayoutNumber: getNextNumber(layoutFiles, /slideLayout(\d+)\.xml/i),
    nextMasterNumber: getNextNumber(masterFiles, /slideMaster(\d+)\.xml/i),
    nextThemeNumber: getNextNumber(themeFiles, /theme(\d+)\.xml/i),
    nextMasterId,
    mediaCounter: getInitialMediaCounter(zip),
    layoutMap: new Map(),
    masterMap: new Map(),
    themeMap: new Map(),
    mediaMap: new Map(),
    newSlides: [],
    newMasters: new Map(),
    newMastersList: [],
    options,
    existingSlideIds,
    existingRelIds,
    existingMasterIds,
  }
}

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

function copyTemplateSlide ({ templateZip, outputZip, templateSlideNumber, position, state, options }) {
  const templateSlidePath = `ppt/slides/slide${templateSlideNumber}.xml`
  const templateSlide = templateZip.file(templateSlidePath)
  if (!templateSlide) {
    console.warn(`⚠️ 模板缺少第${templateSlideNumber}页，跳过注入`)
    return
  }
  const newSlideNumber = state.nextSlideNumber++
  const newSlidePath = `ppt/slides/slide${newSlideNumber}.xml`
  let slideXml = templateSlide.asText()
  if (position === 'start') {
    slideXml = applyCoverPlaceholders(slideXml, options.employee, options.date)
  }
  outputZip.file(newSlidePath, slideXml)
  cloneSlideRelationships({
    templateZip,
    outputZip,
    templateSlideNumber,
    newSlideNumber,
    state,
  })
  ensureContentType(state, `/ppt/slides/slide${newSlideNumber}.xml`, CONTENT_TYPES.slide)
  state.newSlides.push({
    slideNumber: newSlideNumber,
    position,
  })
}

function cloneSlideRelationships ({ templateZip, outputZip, templateSlideNumber, newSlideNumber, state }) {
  const relsPath = `ppt/slides/_rels/slide${templateSlideNumber}.xml.rels`
  const templateRels = templateZip.file(relsPath)
  const newRelsPath = `ppt/slides/_rels/slide${newSlideNumber}.xml.rels`
  if (!templateRels) {
    outputZip.file(newRelsPath, buildEmptyRelationships())
    return
  }
  const relItems = parseRelationships(templateRels.asText())
  const updated = relItems.map((rel) => {
    if (rel.Type === REL_TYPES.slideLayout) {
      const absolutePath = resolveRelationshipPath(`ppt/slides/slide${templateSlideNumber}.xml`, rel.Target)
      const layoutInfo = ensureLayoutAssets({
        templateZip,
        outputZip,
        templateLayoutPath: absolutePath,
        state,
      })
      rel.Target = relativePath(`ppt/slides/slide${newSlideNumber}.xml`, layoutInfo.newPath)
      return rel
    }
    if (rel.Type === REL_TYPES.image) {
      const mapKey = `slide${newSlideNumber}_main`
      if (state.mediaMap.has(mapKey)) {
        const newMediaTarget = state.mediaMap.get(mapKey)
        rel.Target = relativePath(`ppt/slides/slide${newSlideNumber}.xml`, newMediaTarget)
        return rel
      }
      const absoluteMediaPath = resolveRelationshipPath(`ppt/slides/slide${templateSlideNumber}.xml`, rel.Target)
      const newMediaTarget = copyMediaFile({
        templateZip,
        outputZip,
        sourcePath: absoluteMediaPath,
        state,
      })
      rel.Target = relativePath(`ppt/slides/slide${newSlideNumber}.xml`, newMediaTarget)
      return rel
    }
    const absoluteTarget = resolveRelationshipPath(`ppt/slides/slide${templateSlideNumber}.xml`, rel.Target)
    const newTarget = copyGenericPart({
      templateZip,
      outputZip,
      sourcePath: absoluteTarget,
      state,
    })
    if (newTarget) {
      rel.Target = relativePath(`ppt/slides/slide${newSlideNumber}.xml`, newTarget)
    }
    return rel
  })
  outputZip.file(newRelsPath, buildRelationshipsXml(updated))
}

function ensureLayoutAssets ({ templateZip, outputZip, templateLayoutPath, state }) {
  if (state.layoutMap.has(templateLayoutPath)) {
    return state.layoutMap.get(templateLayoutPath)
  }
  const relsPath = layoutRelationshipsPath(templateLayoutPath)
  const relFile = templateZip.file(relsPath)
  if (!relFile) {
    throw new Error(`模板布局缺少关系文件：${relsPath}`)
  }
  const relItems = parseRelationships(relFile.asText())
  const masterRel = relItems.find((rel) => rel.Type === REL_TYPES.slideMaster)
  if (!masterRel) {
    throw new Error(`布局文件未关联母版：${templateLayoutPath}`)
  }
  const masterPath = resolveRelationshipPath(templateLayoutPath, masterRel.Target)
  ensureMasterAssets({
    templateZip,
    outputZip,
    templateMasterPath: masterPath,
    state,
  })
  if (!state.layoutMap.has(templateLayoutPath)) {
    throw new Error(`复制布局失败：${templateLayoutPath}`)
  }
  return state.layoutMap.get(templateLayoutPath)
}

function ensureMasterAssets ({ templateZip, outputZip, templateMasterPath, state }) {
  if (state.masterMap.has(templateMasterPath)) {
    return state.masterMap.get(templateMasterPath)
  }
  const masterFile = templateZip.file(templateMasterPath)
  if (!masterFile) {
    throw new Error(`模板缺少母版文件：${templateMasterPath}`)
  }
  const newMasterNumber = state.nextMasterNumber++
  const newMasterPath = `ppt/slideMasters/slideMaster${newMasterNumber}.xml`
  outputZip.file(newMasterPath, masterFile.asText())
  ensureContentType(state, `/${newMasterPath}`, CONTENT_TYPES.slideMaster)
  const masterRelsPath = masterRelationshipsPath(templateMasterPath)
  const templateRels = templateZip.file(masterRelsPath)
  const newMasterRelsPath = masterRelationshipsPath(newMasterPath)
  let themeInfo = null
  if (templateRels) {
    const relItems = parseRelationships(templateRels.asText())
    const updated = relItems.map((rel) => {
      if (rel.Type === REL_TYPES.theme) {
        const themePath = resolveRelationshipPath(templateMasterPath, rel.Target)
        themeInfo = ensureThemeAssets({
          templateZip,
          outputZip,
          templateThemePath: themePath,
          state,
        })
        rel.Target = relativePath(newMasterPath, themeInfo.newPath)
        return rel
      }
      if (rel.Type === REL_TYPES.slideLayout) {
        const layoutPath = resolveRelationshipPath(templateMasterPath, rel.Target)
        const layoutInfo = copyLayoutFromMaster({
          templateZip,
          outputZip,
          templateLayoutPath: layoutPath,
          newMasterPath,
          state,
        })
        rel.Target = relativePath(newMasterPath, layoutInfo.newPath)
        return rel
      }
      const absolute = resolveRelationshipPath(templateMasterPath, rel.Target)
      const newTarget = copyGenericPart({
        templateZip,
        outputZip,
        sourcePath: absolute,
        state,
      })
      rel.Target = relativePath(newMasterPath, newTarget)
      return rel
    })
    outputZip.file(newMasterRelsPath, buildRelationshipsXml(updated))
  } else {
    outputZip.file(newMasterRelsPath, buildEmptyRelationships())
  }
  const info = {
    newPath: newMasterPath,
    themePath: themeInfo ? themeInfo.newPath : null,
  }
  state.masterMap.set(templateMasterPath, info)
  state.newMasters.set(templateMasterPath, info)
  state.newMastersList.push(info)
  return info
}

function copyLayoutFromMaster ({ templateZip, outputZip, templateLayoutPath, newMasterPath, state }) {
  if (state.layoutMap.has(templateLayoutPath)) {
    return state.layoutMap.get(templateLayoutPath)
  }
  const layoutFile = templateZip.file(templateLayoutPath)
  if (!layoutFile) {
    throw new Error(`模板缺少布局文件：${templateLayoutPath}`)
  }
  const newLayoutNumber = state.nextLayoutNumber++
  const newLayoutPath = `ppt/slideLayouts/slideLayout${newLayoutNumber}.xml`
  outputZip.file(newLayoutPath, layoutFile.asText())
  ensureContentType(state, `/${newLayoutPath}`, CONTENT_TYPES.slideLayout)
  const relsPath = layoutRelationshipsPath(templateLayoutPath)
  const relFile = templateZip.file(relsPath)
  const newRelsPath = layoutRelationshipsPath(newLayoutPath)
  if (relFile) {
    const relItems = parseRelationships(relFile.asText())
    const updated = relItems.map((rel) => {
      if (rel.Type === REL_TYPES.slideMaster) {
        rel.Target = relativePath(newLayoutPath, newMasterPath)
        return rel
      }
      if (rel.Type === REL_TYPES.image) {
        const mediaPath = resolveRelationshipPath(templateLayoutPath, rel.Target)
        const newMediaTarget = copyMediaFile({
          templateZip,
          outputZip,
          sourcePath: mediaPath,
          state,
        })
        rel.Target = relativePath(newLayoutPath, newMediaTarget)
        return rel
      }
      const absolute = resolveRelationshipPath(templateLayoutPath, rel.Target)
      const newTarget = copyGenericPart({
        templateZip,
        outputZip,
        sourcePath: absolute,
        state,
      })
      rel.Target = relativePath(newLayoutPath, newTarget)
      return rel
    })
    outputZip.file(newRelsPath, buildRelationshipsXml(updated))
  } else {
    outputZip.file(newRelsPath, buildEmptyRelationships())
  }
  const info = { newPath: newLayoutPath }
  state.layoutMap.set(templateLayoutPath, info)
  return info
}

function ensureThemeAssets ({ templateZip, outputZip, templateThemePath, state }) {
  if (state.themeMap.has(templateThemePath)) {
    return state.themeMap.get(templateThemePath)
  }
  const themeFile = templateZip.file(templateThemePath)
  if (!themeFile) {
    throw new Error(`模板缺少主题文件：${templateThemePath}`)
  }
  const newThemeNumber = state.nextThemeNumber++
  const newThemePath = `ppt/theme/theme${newThemeNumber}.xml`
  outputZip.file(newThemePath, themeFile.asText())
  ensureContentType(state, `/${newThemePath}`, CONTENT_TYPES.theme)
  const info = { newPath: newThemePath }
  state.themeMap.set(templateThemePath, info)
  return info
}

function copyMediaFile ({ templateZip, outputZip, sourcePath, state }) {
  if (state.mediaMap.has(sourcePath)) {
    return state.mediaMap.get(sourcePath)
  }
  const file = templateZip.file(sourcePath)
  if (!file) {
    // 修复：如果模板中引用了不存在的图片，不要抛出错误，而是跳过，避免破坏流程
    console.warn(`⚠️ 模板中引用了不存在的资源：${sourcePath}，已跳过复制。`)
    return sourcePath // 返回原路径，尽管它可能无效
  }
  const ext = path.extname(sourcePath) || '.bin'
  const newName = `ppt/media/template_media_${state.mediaCounter++}${ext}`
  outputZip.file(newName, file.asNodeBuffer())
  state.mediaMap.set(sourcePath, newName)
  return newName
}

function copyGenericPart ({ templateZip, outputZip, sourcePath, state }) {
  if (!sourcePath.startsWith('ppt/')) {
    return sourcePath
  }
  if (outputZip.file(sourcePath)) {
    return sourcePath
  }
  const part = templateZip.file(sourcePath)
  if (part) {
    if (part.options && part.options.binary) {
      outputZip.file(sourcePath, part.asNodeBuffer())
    } else {
      outputZip.file(sourcePath, part.asText())
    }
  }
  if (sourcePath.startsWith('ppt/notesSlides/')) {
    ensureContentType(
      state,
      `/${sourcePath}`,
      'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    )
  }
  return sourcePath
}

// --- 核心修复：Content Type 检查逻辑 ---
function ensureDefaultContentType (state, extension, contentType) {
  const ext = extension.replace(/^\./, '').toLowerCase()

  // 修复：使用正则检查 Extension 属性，不限定顺序，支持单双引号
  // 匹配 <Default ... Extension="png" ...> 或 <Default ... Extension='png' ...>
  const regex = new RegExp(`<Default[^>]*Extension=["']${ext}["'][^>]*>`, 'i')

  if (state.contentTypesContent.match(regex)) {
    return
  }

  const defaultEntry = `<Default Extension="${ext}" ContentType="${contentType}"/>`
  const typesStartMatch = state.contentTypesContent.match(/<Types[^>]*>/)
  if (!typesStartMatch) return
  const insertIndex = typesStartMatch.index + typesStartMatch[0].length
  state.contentTypesContent =
    state.contentTypesContent.slice(0, insertIndex) +
    defaultEntry +
    state.contentTypesContent.slice(insertIndex)
}

function ensureContentType (state, partName, contentType) {
  // 修复：PartName 匹配支持单双引号
  const partNameRegex = new RegExp(`<Override[^>]*PartName=["']${partName}["'][^>]*>`, 'i')
  if (state.contentTypesContent.match(partNameRegex)) {
    return
  }
  const override = `  <Override PartName="${partName}" ContentType="${contentType}"/>`
  const typesEndIndex = state.contentTypesContent.lastIndexOf('</Types>')
  if (typesEndIndex === -1) {
    throw new Error('Content_Types.xml缺少</Types>结束标签')
  }
  state.contentTypesContent = state.contentTypesContent.slice(0, typesEndIndex) + `\n${override}\n` + state.contentTypesContent.slice(typesEndIndex)
}

function updatePresentationDocuments (zip, state) {
  if (!state.newSlides.length) {
    return
  }
  const slideEntries = extractExistingSlideEntries(state.presContent)
  const updatedSlideEntries = buildSlideEntries(slideEntries, state)
  state.presContent = state.presContent.replace(slideEntries.raw, updatedSlideEntries.xml)
  const updatedRels = appendSlideRelationships(state.presRelsContent, state)
  state.presRelsContent = updatedRels
  const updatedMasters = appendMasterEntries(state.presContent, state)
  state.presContent = updatedMasters
  for (const slide of state.newSlides) {
    ensureContentType(state, `/ppt/slides/slide${slide.slideNumber}.xml`, CONTENT_TYPES.slide)
  }
  for (const [templatePath, layoutInfo] of state.layoutMap.entries()) {
    ensureContentType(state, `/${layoutInfo.newPath}`, CONTENT_TYPES.slideLayout)
  }
  for (const [templatePath, masterInfo] of state.masterMap.entries()) {
    ensureContentType(state, `/${masterInfo.newPath}`, CONTENT_TYPES.slideMaster)
  }
  for (const [templatePath, themeInfo] of state.themeMap.entries()) {
    ensureContentType(state, `/${themeInfo.newPath}`, CONTENT_TYPES.theme)
  }
  zip.file('ppt/presentation.xml', state.presContent)
  zip.file('ppt/_rels/presentation.xml.rels', state.presRelsContent)
  zip.file('[Content_Types].xml', state.contentTypesContent)
  state.newSlides = []
}

function extractExistingSlideEntries (presContent) {
  const match = presContent.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/) || { raw: '<p:sldIdLst></p:sldIdLst>', inner: '' }
  return { raw: match[0], inner: match[1] }
}

function buildSlideEntries (existing, state) {
  const existingEntries = existing.inner
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
  const newSlides = state.newSlides.map((slide) => {
    while (state.existingSlideIds.has(state.nextSlideId) || state.nextSlideId <= 256) {
      state.nextSlideId++
    }
    slide.slideId = state.nextSlideId
    state.existingSlideIds.add(state.nextSlideId)
    state.nextSlideId++
    while (state.existingRelIds.has(state.nextRelId)) {
      state.nextRelId++
    }
    slide.relId = `rId${state.nextRelId}`
    state.existingRelIds.add(state.nextRelId)
    state.nextRelId++
    return slide
  })
  const entryFor = (slide) => `    <p:sldId id="${slide.slideId}" r:id="${slide.relId}"/>`
  const startEntries = newSlides.filter((s) => s.position === 'start').map(entryFor)
  const endEntries = newSlides.filter((s) => s.position === 'end').map(entryFor)
  const xml = `<p:sldIdLst>
${[...startEntries, ...existingEntries, ...endEntries].join('\n')}
</p:sldIdLst>`
  return { xml }
}

function appendSlideRelationships (relsContent, state) {
  const insertIndex = relsContent.indexOf('</Relationships>')
  if (insertIndex === -1) {
    throw new Error('presentation.xml.rels 内容异常')
  }
  const slideRels = state.newSlides
    .map(
      (slide) =>
        `  <Relationship Id="${slide.relId}" Type="${REL_TYPES.slide}" Target="slides/slide${slide.slideNumber}.xml"/>`,
    )
    .join('\n')
  return `${relsContent.slice(0, insertIndex)}${slideRels}\n${relsContent.slice(insertIndex)}`
}

function appendMasterEntries (presContent, state) {
  if (!state.newMastersList.length) {
    return presContent
  }
  const relsContent = state.presRelsContent
  const insertIndex = relsContent.indexOf('</Relationships>')
  if (insertIndex === -1) {
    throw new Error('presentation.xml.rels 缺少 Relationships 结束标签')
  }
  const masterListMatch = presContent.match(/<p:sldMasterIdLst>([\s\S]*?)<\/p:sldMasterIdLst>/) || { 1: '' }
  const existingEntries = masterListMatch[1]
    ? masterListMatch[1]
      .split('\n')
      .map((line) => line.trim())
      .filter(Boolean)
    : []
  const newMasterEntries = []
  const newMasterRels = []
  for (const info of state.newMastersList) {
    while (state.existingMasterIds.has(state.nextMasterId)) {
      state.nextMasterId++
    }
    while (state.existingRelIds.has(state.nextRelId)) {
      state.nextRelId++
    }
    const relId = `rId${state.nextRelId}`
    const entry = `    <p:sldMasterId id="${state.nextMasterId}" r:id="${relId}"/>`
    newMasterEntries.push(entry)
    newMasterRels.push(
      `  <Relationship Id="${relId}" Type="${REL_TYPES.slideMaster}" Target="${info.newPath.replace(
        'ppt/',
        '',
      )}"/>`,
    )
    state.existingMasterIds.add(state.nextMasterId)
    state.existingRelIds.add(state.nextRelId)
    state.nextMasterId++
    state.nextRelId++
  }
  const newMasterList = `<p:sldMasterIdLst>
${[...existingEntries, ...newMasterEntries].join('\n')}
</p:sldMasterIdLst>`
  let updatedPres = presContent
  if (masterListMatch[0]) {
    updatedPres = presContent.replace(masterListMatch[0], newMasterList)
  } else {
    updatedPres = presContent.replace(
      '</p:presentation>',
      `${newMasterList}\n</p:presentation>`,
    )
  }
  const updatedRelsContent = `${relsContent.slice(
    0,
    insertIndex,
  )}${newMasterRels.join('\n')}\n${relsContent.slice(insertIndex)}`
  state.presRelsContent = updatedRelsContent
  return updatedPres
}

function parseRelationships (xml) {
  const relRegex = /<Relationship\s+([^>]+?)\/>/g
  const attrRegex = /([A-Za-z:]+)\s*=\s*["']([^"']*)["']/g
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

function buildEmptyRelationships () {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`
}

function resolveRelationshipPath (from, target) {
  const baseDir = path.posix.dirname(from)
  return path.posix.normalize(path.posix.join(baseDir, target))
}

function relativePath (from, to) {
  const fromDir = path.posix.dirname(from)
  let relPath = path.posix.relative(fromDir, to)
  relPath = relPath.replace(/\\/g, '/')
  relPath = relPath.replace(/^\.\//g, '')
  return relPath
}

function layoutRelationshipsPath (layoutPath) {
  const baseName = path.posix.basename(layoutPath)
  return `ppt/slideLayouts/_rels/${baseName}.rels`
}

function masterRelationshipsPath (masterPath) {
  const baseName = path.posix.basename(masterPath)
  return `ppt/slideMasters/_rels/${baseName}.rels`
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

function applyCoverPlaceholders (xml, employee = {}, dateValue = new Date()) {
  const dateText = formatCnDate(dateValue)
  const replacements = {
    姓名: employee.name || '',
    性别: employee.gender || '',
    年龄: employee.age || '',
    日期: dateText || '',
    工号: employee.id || '',
  }
  return replaceTextInSlidePreferred(xml, replacements)
}

function formatCnDate (dateInput) {
  const date = dateInput instanceof Date ? dateInput : new Date(dateInput || Date.now())
  if (Number.isNaN(date.getTime())) {
    return ''
  }
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}年${month}月${day}日`
}

function escapeXmlValue (value) {
  if (!value) return ''
  return String(value)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

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
      const placeholderPattern = new RegExp(`\\{\\{${optionalSpaces}${keyPattern}${optionalSpaces}\\}\}${optionalSpaces}`, 'gi')
      updatedContent = updatedContent.replace(placeholderPattern, (phMatch) => {
        const suffixMatch = phMatch.match(/((?:\\s|\\u00A0|&nbsp;|&#160;|&#xA0;)*)$/i)
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

function escapeRegex (str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function replaceImagePlaceholder (slideXml, image, state, outputZip, slideNumber) {
  try {
    let ext = '.png'
    let mimeType = 'image/png'
    if (image.data) {
      const dataUrlMatch = image.data.match(/^data:(image\/\w+);base64,/)
      if (dataUrlMatch) {
        mimeType = dataUrlMatch[1]
        ext = `.${mimeType.split('/')[1]}`
      }
    } else if (image.fullPath) {
      ext = path.extname(image.fullPath) || '.png'
      mimeType = ext.toLowerCase().includes('png') ? 'image/png' :
        ext.toLowerCase().includes('jpg') || ext.toLowerCase().includes('jpeg') ? 'image/jpeg' :
          ext.toLowerCase().includes('gif') ? 'image/gif' :
            ext.toLowerCase().includes('bmp') ? 'image/bmp' : 'image/png'
    } else {
      console.warn(`⚠️ 图片数据为空，跳过处理`)
      return slideXml
    }
    const newImagePath = `ppt/media/image_${state.mediaCounter++}${ext}`
    ensureDefaultContentType(state, ext, mimeType)
    let buffer
    if (image.data) {
      const base64Data = image.data.replace(/^data:image\/\w+;base64,/, '')
      buffer = Buffer.from(base64Data, 'base64')
    } else if (image.fullPath) {
      buffer = fs.readFileSync(image.fullPath)
    } else {
      console.warn(`⚠️ 图片数据为空，跳过处理`)
      return slideXml
    }
    if (outputZip) {
      outputZip.file(newImagePath, buffer)
    }
    const mapKey = `slide${slideNumber}_main`
    state.mediaMap.set(mapKey, newImagePath)
    let updatedXml = slideXml
    updatedXml = updatedXml.replace(/\{\{\s*图片\s*\}\}/g, '')
    return updatedXml
  } catch (err) {
    console.warn(`⚠️ 替换幻灯片图片失败：${err.message}`)
    return slideXml
  }
}

async function copyTemplateSecondPageForImages (templatePath, outputPath, imageItems, employee) {
  const [templateBuffer, outputBuffer] = await Promise.all([
    fs.readFile(templatePath),
    fs.readFile(outputPath),
  ])
  const templateZip = new PizZip(templateBuffer)
  const outputZip = new PizZip(outputBuffer)
  const state = initializeState(outputZip, { employee, date: new Date() })
  const newSlides = []
  for (const image of imageItems) {
    const cat = image.category || 'other'
    const templateSlideNumber = cat === 'inbody' ? 2 : cat === 'lab' ? 3 : cat === 'ecg' ? 4 : cat === 'ai' ? 5 : 3
    const templateSlidePath = `ppt/slides/slide${templateSlideNumber}.xml`
    const templateSlide = templateZip.file(templateSlidePath)
    if (!templateSlide) {
      console.warn(`⚠️ 模板缺少第${templateSlideNumber}页，跳过注入`)
      continue
    }
    const newSlideNumber = state.nextSlideNumber++
    const newSlidePath = `ppt/slides/slide${newSlideNumber}.xml`
    let slideXml = templateSlide.asText()
    const replacements = {
      姓名: employee.name || '',
      工号: employee.id || '',
    }
    slideXml = replaceTextInSlidePreferred(slideXml, replacements)
    slideXml = replaceImagePlaceholder(slideXml, image, state, outputZip, newSlideNumber)
    outputZip.file(newSlidePath, slideXml)
    cloneSlideRelationships({
      templateZip,
      outputZip,
      templateSlideNumber,
      newSlideNumber,
      state,
    })
    ensureContentType(state, `/ppt/slides/slide${newSlideNumber}.xml`, CONTENT_TYPES.slide)
    newSlides.push({
      slideNumber: newSlideNumber,
      position: 'start',
    })
  }
  if (!newSlides.length) {
    return
  }
  state.newSlides = newSlides
  updatePresentationDocuments(outputZip, state)
  const updatedBuffer = outputZip.generate({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
  await fs.writeFile(outputPath, updatedBuffer)
}

async function main () {
  await fs.ensureDir(OUTPUT_DIR)
  await fs.ensureDir(PDF_IMAGE_DIR)
  await fs.ensureDir(PDF_TEMP_DIR)
  const templatePath = await resolveExistingPath(TEMPLATE_CANDIDATES, '模板文件')
  const sheetPath = await resolveExistingPath(EMPLOYEE_SHEET_CANDIDATES, '员工表')
  const layout = await loadTemplateLayout(templatePath)
  const employees = await loadEmployees(sheetPath)
  if (!employees.length) {
    console.warn('⚠️ 员工表为空，已结束。')
    return
  }
  const availableFiles = await safeReadDir(DATA_DIR)
  const successReports = []
  const skippedEmployees = []
  const nameCounter = new Map()
  for (const employee of employees) {
    try {
      const assetInfo = await collectEmployeeAssets(employee, availableFiles)
      const hasSummary = Boolean(assetInfo.summaryText.trim())
      const hasAttachments = assetInfo.attachments.length > 0
      if (!hasSummary && !hasAttachments) {
        skippedEmployees.push({
          employee,
          reason: '缺少体检结果与AI总结',
        })
        continue
      }
      const pptx = initializePresentation(layout)
      const currentCount = nameCounter.get(employee.name) || 0
      const suffix = currentCount > 0 ? `_${currentCount}` : ''
      const outputName = buildReportFileName(employee, suffix)
      const outputPath = path.join(OUTPUT_DIR, outputName)
      nameCounter.set(employee.name, currentCount + 1)
      await pptx.writeFile({ fileName: outputPath })
      const imageItems = await buildImageItems(assetInfo, employee)
      if (imageItems.length > 0) {
        await copyTemplateSecondPageForImages(templatePath, outputPath, imageItems, employee)
      }
      await insertTemplateSlides(templatePath, outputPath, { employee, date: new Date(), summary: assetInfo.summaryText })
      successReports.push({ employee, outputPath })
      console.log(`✓ 已生成 ${employee.name}（${employee.id}）：${outputPath}`)
    } catch (error) {
      skippedEmployees.push({
        employee,
        reason: `生成失败：${error.message}`,
      })
      console.error(`❌ ${employee.name} 生成失败：${error.message}`)
    }
  }
  console.log('\n===== 生成统计 =====')
  console.log(`✅ 已生成：${successReports.length} 人`)
  successReports.forEach((item) => {
    console.log(`  - ${item.employee.name}（${item.employee.id}） -> ${item.outputPath}`)
  })
  console.log(`⚠️ 未生成：${skippedEmployees.length} 人`)
  skippedEmployees.forEach((item) => {
    console.log(`  - ${item.employee.name}（${item.employee.id}）：${item.reason}`)
  })
}

main().catch((error) => {
  console.error('❌ 生成失败：', error)
  process.exitCode = 1
})