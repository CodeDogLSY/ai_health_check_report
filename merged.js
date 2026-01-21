/**
 * 员工体检报告批量生成脚本
 * - 从「员工表.xlsx」读取员工信息
 * - 在 data/ 目录匹配员工附件（图片、PDF、AI总结docx）
 * - 基于 PPTX 模板插入：封面 -> 附件页(按类别) -> 指导页(可选) -> 结束页
 * - 可选：将生成的 PPTX 转成 PDF，并在成功后删除 PPTX
 *
 * CLI:
 *   node merged.js                 # 主流程：生成报告
 *   node merged.js --convert-ppt-pdf      # 转换 result_add_suggest/ 下所有 PPTX 为 PDF
 *   node merged.js --rename-ppt-files     # 给 result_add_suggest/ 下 PPTX 文件名补充证件号
 *   node merged.js --send-wechat-ee-msg  # 读取 send_data/ 下 PDF，提取证件号并尝试发送企信
 *
 */

const path = require("path")
const fs = require("fs-extra")
const ExcelJS = require("exceljs")
const PizZip = require("pizzip")
const PptxGenJS = require("pptxgenjs")
const mammoth = require("mammoth")
const { fromPath: pdfFromPath } = require("pdf2pic")
const { createCanvas } = require("canvas")
const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.js")
const crypto = require("crypto")
const libre = require("libreoffice-convert")
const util = require("util")
const convertAsync = util.promisify(libre.convert)
const { spawn } = require("child_process")
// --------------------------------------------

pdfjsLib.GlobalWorkerOptions.workerSrc =
  require.resolve("pdfjs-dist/legacy/build/pdf.worker.js")

//是否带第六页指导页
const has6WhitePage = false
//PDF 生成文件类型开关 设置为 true 开启 PDF 转换，false 则关闭
const generatePdf = true
//name中加入id开关 false 不添加 true 添加
const addId = true

const ROOT = path.resolve(__dirname, ".")
const DATA_DIR = path.join(ROOT, "data")
const OUTPUT_DIR = path.join(ROOT, "output")
const PDF_IMAGE_DIR = path.join(OUTPUT_DIR, "_pdf_images")
const PDF_TEMP_DIR = path.join(PDF_IMAGE_DIR, "_tmp")

// 并发控制配置
const MAX_CONCURRENT_EMPLOYEES = 6 // 同时处理的员工数量
const MAX_CONCURRENT_PDF_CONVERSIONS = 1 // 同时进行的 PDF 转换数量（Windows上建议设为1以确保稳定性）

const TEMPLATE_CANDIDATES = ["2025员工体检报告（模板）.pptx", "template.pptx"]
const EMPLOYEE_SHEET_CANDIDATES = ["员工表.xlsx", "employees.xlsx"]

const IMAGE_EXTENSIONS = new Set([".png", ".jpg", ".jpeg", ".gif", ".bmp"])
const PDF_EXTENSIONS = new Set([".pdf"])
const EMU_PER_INCH = 914400

const DEFAULT_LAYOUT = {
  width: 10,
  height: 5.625,
  orientation: "landscape",
}

function normalizeName (value, id = "") {
  if (!value) return ""
  const namePart = String(value)
    .replace(/\.[^.]+$/, "")
    .replace(/\s+/g, "")
    .toLowerCase()
  if (id) {
    return namePart + id.toLowerCase().replace(/\s+/g, "")
  }
  const match = namePart.match(/^([^-]+)-([^-]+)/)
  if (match) {
    return match[1] + match[2]
  }
  return namePart.replace(/[-_].*$/, "")
}

function normalizeText (value) {
  if (value === undefined || value === null) return ""
  return String(value).trim()
}

function sanitizeForFilename (value) {
  return value.replace(/[<>:"/\\|?*]/g, "").replace(/\s+/g, "")
}

function buildAsciiSafeLabel (value) {
  const sanitized = sanitizeForFilename(value).replace(/[^\x00-\x7F]/g, "")
  if (sanitized) {
    return sanitized
  }
  return `pdf_${Date.now()}`
}

function formatDate (date) {
  const hh = String(date.getHours()).padStart(2, "0")
  const mi = String(date.getMinutes()).padStart(2, "0")
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
  throw new Error(
    `${label}缺失，请确认以下文件之一存在：${candidates.join(", ")}`,
  )
}

async function loadTemplateLayout (templatePath) {
  const fallback = { ...DEFAULT_LAYOUT }
  if (!(await fs.pathExists(templatePath))) {
    return fallback
  }
  try {
    const buffer = await fs.readFile(templatePath)
    const zip = new PizZip(buffer)
    const presentationFile = zip.file("ppt/presentation.xml")
    if (!presentationFile) {
      return fallback
    }
    const xml = presentationFile.asText()
    const sizeMatch = xml.match(
      /<p:sldSz[^>]*cx=["'](\d+)["'][^>]*cy=["'](\d+)["']/,
    )
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
    const orientation = width >= height ? "landscape" : "portrait"
    return { width, height, orientation }
  } catch (error) {
    console.warn(
      `⚠️ 模板布局解析失败，将使用默认 16:9 布局。原因：${error.message}`,
    )
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
  const headers = headerRow.values.map((value) =>
    typeof value === "string" ? value.trim() : value,
  )
  const employees = []
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return
    const record = {}
    row.values.forEach((cell, cellIndex) => {
      const header = headers[cellIndex]
      if (!header || typeof header !== "string") return
      record[header] = typeof cell === "string" ? cell.trim() : cell
    })
    const name = normalizeText(record["姓名"] || record["name"])
    if (!name) return
    employees.push({
      name,
      gender: normalizeText(record["性别"]) || "未知",
      age: normalizeText(record["年龄"]) || "未知",
      id: normalizeText(record["证件号"]) || "未知",
      date: normalizeText(record["体检日期"]) || "未知",
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

/**
 * 根据文件名关键字为附件分配“显示顺序优先级”（数值越小越靠前）。
 * 规则保持与原逻辑一致，仅抽成函数方便维护。
 */
function getAttachmentPriority (lowerFileName) {
  if (lowerFileName.includes("inbody")) return 1
  if (lowerFileName.includes("尿常规") || lowerFileName.includes("尿检"))
    return 2
  if (lowerFileName.includes("血常规")) return 3
  if (lowerFileName.includes("生化检查")) return 4
  if (lowerFileName.includes("糖化血红蛋白")) return 5
  if (lowerFileName.includes("心电图")) return 6
  if (lowerFileName.includes("ai解读")) return 7
  return 99
}

/**
 * 根据文件名关键字给附件分类（用于匹配模板页：inbody / lab / ecg / ai）。
 * 规则保持与原逻辑一致。
 */
function getAttachmentCategory (lowerFileName) {
  if (lowerFileName.includes("inbody")) return "inbody"
  if (lowerFileName.includes("心电图")) return "ecg"
  if (
    lowerFileName.includes("尿常规") ||
    lowerFileName.includes("尿检") ||
    lowerFileName.includes("血常规") ||
    lowerFileName.includes("生化检查") ||
    lowerFileName.includes("糖化血红蛋白")
  ) {
    return "lab"
  }
  if (lowerFileName.includes("ai解读")) return "ai"
  return "other"
}

async function collectEmployeeAssets (employee, files) {
  const normalized = normalizeName(employee.name, employee.id)
  const matchingFiles = files.filter(
    (file) => normalizeName(file) === normalized,
  )
  const summaryFile = matchingFiles.find(
    (file) =>
      file.includes("总结") && path.extname(file).toLowerCase() === ".docx",
  )
  const summaryText = summaryFile
    ? await extractSummaryText(path.join(DATA_DIR, summaryFile))
    : ""
  const attachments = matchingFiles
    .filter((file) => file !== summaryFile)
    .map((file) => {
      const fullPath = path.join(DATA_DIR, file)
      const ext = path.extname(file).toLowerCase()
      const type = IMAGE_EXTENSIONS.has(ext)
        ? "image"
        : PDF_EXTENSIONS.has(ext)
          ? "pdf"
          : "other"
      const label = buildAttachmentLabel(file)
      const lower = file.toLowerCase()

      return {
        fileName: file,
        fullPath,
        label,
        type,
        ext,
        priority: getAttachmentPriority(lower),
        category: getAttachmentCategory(lower),
      }
    })
    .sort((a, b) => a.priority - b.priority)
  return {
    employee,
    summaryFile: summaryFile ? path.join(DATA_DIR, summaryFile) : null,
    summaryText: summaryText || "",
    attachments,
    totalFiles: matchingFiles.length,
  }
}

async function extractSummaryText (filePath) {
  try {
    const { value } = await mammoth.extractRawText({ path: filePath })
    return value
      .split("\n")
      .map((item) => item.trim())
      .filter(Boolean)
      .join("\n")
  } catch (error) {
    console.warn(`⚠️ 无法读取AI总结：${filePath}，原因：${error.message}`)
    return ""
  }
}

function buildAttachmentLabel (fileName) {
  const base = path.basename(fileName, path.extname(fileName))
  const parts = base.split(/[-_]/)
  return parts.slice(1).join("-") || "附件"
}

async function buildImageItems (assetInfo, employee) {
  const imageItems = []
  const pdfTasks = []

  // 先收集所有图片
  for (const attachment of assetInfo.attachments) {
    if (attachment.type === "image") {
      imageItems.push({
        label: attachment.label,
        fullPath: attachment.fullPath,
        category: attachment.category,
      })
    } else if (attachment.type === "pdf") {
      pdfTasks.push(() => convertPdfAttachment(attachment, employee))
    }
  }

  // 并行处理所有 PDF 转换
  if (pdfTasks.length > 0) {
    const convertedResults = await Promise.all(pdfTasks.map((task) => task()))
    for (const converted of convertedResults) {
      imageItems.push(...converted)
    }
  }

  return imageItems
}

async function convertPdfAttachment (pdfAttachment, employee) {
  let tempPdfPath = ""
  try {
    const safeLabel = buildAsciiSafeLabel(
      `${employee.id}_${employee.name}_${pdfAttachment.label}`,
    )
    const tempPdfName = `${crypto.randomUUID()}.pdf`
    tempPdfPath = path.join(PDF_TEMP_DIR, tempPdfName)
    await fs.ensureDir(PDF_TEMP_DIR)
    await fs.copyFile(pdfAttachment.fullPath, tempPdfPath)
    const pagesFromPic = await convertWithPdf2Pic(
      tempPdfPath,
      pdfAttachment,
      safeLabel,
    )
    if (pagesFromPic.length) {
      return pagesFromPic.map((p) => ({
        ...p,
        category: pdfAttachment.category,
      }))
    }
    console.warn(`⚠️ pdf2pic 转换未返回结果，尝试使用 pdfjs-dist 兜底...`)
    const pdfImgPages = await convertWithPdfRenderer(
      tempPdfPath,
      pdfAttachment,
    )
    if (pdfImgPages.length) {
      return pdfImgPages.map((p) => ({
        ...p,
        category: pdfAttachment.category,
      }))
    }
    return []
  } catch (error) {
    console.warn(
      `⚠️ PDF 转图失败（${pdfAttachment.fileName}）：${error.message}`,
    )
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
      cMapUrl: "./node_modules/pdfjs-dist/cmaps/",
      cMapPacked: true,
    }).promise

    // 并行处理所有页面
    const pageTasks = []
    for (let pageNumber = 1; pageNumber <= pdfDoc.numPages; pageNumber++) {
      pageTasks.push(async () => {
        const page = await pdfDoc.getPage(pageNumber)
        const viewport = page.getViewport({ scale: 10 })
        const canvas = createCanvas(viewport.width, viewport.height)
        const context = canvas.getContext("2d")
        await page.render({ canvasContext: context, viewport }).promise
        const imageBuffer = canvas.toBuffer("image/png")
        return {
          label: `${attachment.label} 第${pageNumber}页`,
          data: `data:image/png;base64,${imageBuffer.toString("base64")}`,
        }
      })
    }

    const pages = await Promise.all(pageTasks.map((task) => task()))
    return pages
  } catch (error) {
    console.warn(
      `⚠️ 内置 PDF 渲染失败（${attachment.fileName}）：${error.message}`,
    )
    return []
  }
}

async function convertWithPdf2Pic (pdfPath, attachment, safeLabel) {
  try {
    await fs.ensureDir(PDF_IMAGE_DIR)
    const convert = pdfFromPath(pdfPath, {
      density: 900,
      format: "png",
      width: 2400,
      height: 1600,
      saveFilename: safeLabel,
      savePath: PDF_IMAGE_DIR,
    })
    const pages = await convert.bulk(-1, { responseType: "base64" })
    return pages.map((pageResult, idx) => ({
      label: `${attachment.label} 第${pageResult.page || idx + 1}页`,
      data: `data:image/png;base64,${pageResult.base64}`,
    }))
  } catch (error) {
    console.warn(
      `⚠️ pdf2pic 转换失败（${attachment.fileName}）：${error.message}`,
    )
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
  pptx.author = "Health Manage AI"
  pptx.company = "Health Manage AI"
  pptx.subject = "员工体检报告"
  pptx.title = "2025 员工体检报告"
  return pptx
}

function buildReportFileName (employee, suffix = "") {
  const safeName = sanitizeForFilename(employee.name)
  const safeId = sanitizeForFilename(employee.id)
  const dateStr = formatDate(new Date())
  const idPart = addId ? `_${safeId}` : ""
  if (suffix) {
    // return `员工体检报告_${safeName}${suffix}_${dateStr}.pptx`
    return `体检报告_${safeName}${idPart}${suffix}.pptx`
  }
  return `体检报告_${safeName}${idPart}.pptx`
}

const CONTENT_TYPES = {
  slide:
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
  slideLayout:
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
  slideMaster:
    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
}

const REL_TYPES = {
  slide:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
  slideLayout:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  slideMaster:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
  theme:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  image:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
}

/**
 * 插入模板中的「单页」到目标 PPTX。
 * 说明：此函数只负责把模板页复制进目标文件，并更新 presentation / rels / content-types。
 * 图片、文字占位符的替换由 copyTemplateSlide 内部处理（封面页）或后续图片页函数处理。
 */
async function insertTemplateSlide ({
  templatePath,
  outputPath,
  templateSlideNumber,
  position,
  options = {},
}) {
  const [templateData, outputBuffer] = await Promise.all([
    getTemplateBuffer(templatePath),
    fs.readFile(outputPath),
  ])
  const templateZip = templateData.zip
  const outputZip = new PizZip(outputBuffer)
  const state = initializeState(outputZip, options)

  copyTemplateSlide({
    templateZip,
    outputZip,
    templateSlideNumber,
    position,
    state,
    options,
  })

  if (!state.newSlides.length) {
    return
  }

  updatePresentationDocuments(outputZip, state)
  const updatedBuffer = outputZip.generate({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  })
  await fs.writeFile(outputPath, updatedBuffer)
}

// 分别插入封面和总结页的函数
async function insertCoverSlide (templatePath, outputPath, options = {}) {
  return insertTemplateSlide({
    templatePath,
    outputPath,
    templateSlideNumber: 1,
    position: "start",
    options,
  })
}

async function insertTextGuidanceSlide (templatePath, outputPath, options = {}) {
  return insertTemplateSlide({
    templatePath,
    outputPath,
    templateSlideNumber: 6,
    position: "middle", // 在图片页之后、结束页之前
    options,
  })
}

async function insertSummarySlides (templatePath, outputPath, options = {}) {
  return insertTemplateSlide({
    templatePath,
    outputPath,
    templateSlideNumber: 7,
    position: "end",
    options,
  })
}

function initializeState (zip, options) {
  const presentationXml = zip.file("ppt/presentation.xml")
  const presentationRels = zip.file("ppt/_rels/presentation.xml.rels")
  const contentTypes = zip.file("[Content_Types].xml")
  if (!presentationXml || !presentationRels || !contentTypes) {
    throw new Error("目标 PPT 结构不完整，缺少 presentation 元数据")
  }
  const presContent = presentationXml.asText()
  const presRelsContent = presentationRels.asText()

  // 修复：使用更宽松的正则匹配 ID，支持单双引号
  const slideIdMatches = [
    ...presContent.matchAll(/<p:sldId[^>]*id=["'](\d+)["']/g),
  ]
  const existingSlideIds = new Set(
    slideIdMatches.map((m) => parseInt(m[1], 10)),
  )

  const relIdMatches = [...presRelsContent.matchAll(/Id=["']rId(\d+)["']/g)]
  const existingRelIds = new Set(relIdMatches.map((m) => parseInt(m[1], 10)))

  const masterIdMatches = [
    ...presContent.matchAll(/<p:sldMasterId[^>]*id=["'](\d+)["']/g),
  ]
  const existingMasterIds = new Set(
    masterIdMatches.map((m) => parseInt(m[1], 10)),
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

function copyTemplateSlide ({
  templateZip,
  outputZip,
  templateSlideNumber,
  position,
  state,
  options,
}) {
  const templateSlidePath = `ppt/slides/slide${templateSlideNumber}.xml`
  const templateSlide = templateZip.file(templateSlidePath)
  if (!templateSlide) {
    console.warn(`⚠️ 模板缺少第${templateSlideNumber}页，跳过注入`)
    return
  }
  const newSlideNumber = state.nextSlideNumber++
  const newSlidePath = `ppt/slides/slide${newSlideNumber}.xml`
  let slideXml = templateSlide.asText()
  if (position === "start") {
    slideXml = applyCoverPlaceholders(slideXml, options.employee)
  }
  outputZip.file(newSlidePath, slideXml)
  cloneSlideRelationships({
    templateZip,
    outputZip,
    templateSlideNumber,
    newSlideNumber,
    state,
  })
  ensureContentType(
    state,
    `/ppt/slides/slide${newSlideNumber}.xml`,
    CONTENT_TYPES.slide,
  )
  state.newSlides.push({
    slideNumber: newSlideNumber,
    position,
  })
}

function cloneSlideRelationships ({
  templateZip,
  outputZip,
  templateSlideNumber,
  newSlideNumber,
  state,
}) {
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
      const absolutePath = resolveRelationshipPath(
        `ppt/slides/slide${templateSlideNumber}.xml`,
        rel.Target,
      )
      const layoutInfo = ensureLayoutAssets({
        templateZip,
        outputZip,
        templateLayoutPath: absolutePath,
        state,
      })
      rel.Target = relativePath(
        `ppt/slides/slide${newSlideNumber}.xml`,
        layoutInfo.newPath,
      )
      return rel
    }
    if (rel.Type === REL_TYPES.image) {
      const mapKey = `slide${newSlideNumber}_main`
      if (state.mediaMap.has(mapKey)) {
        const newMediaTarget = state.mediaMap.get(mapKey)
        rel.Target = relativePath(
          `ppt/slides/slide${newSlideNumber}.xml`,
          newMediaTarget,
        )
        return rel
      }
      const absoluteMediaPath = resolveRelationshipPath(
        `ppt/slides/slide${templateSlideNumber}.xml`,
        rel.Target,
      )
      const newMediaTarget = copyMediaFile({
        templateZip,
        outputZip,
        sourcePath: absoluteMediaPath,
        state,
      })
      rel.Target = relativePath(
        `ppt/slides/slide${newSlideNumber}.xml`,
        newMediaTarget,
      )
      return rel
    }
    const absoluteTarget = resolveRelationshipPath(
      `ppt/slides/slide${templateSlideNumber}.xml`,
      rel.Target,
    )
    const newTarget = copyGenericPart({
      templateZip,
      outputZip,
      sourcePath: absoluteTarget,
      state,
    })
    if (newTarget) {
      rel.Target = relativePath(
        `ppt/slides/slide${newSlideNumber}.xml`,
        newTarget,
      )
    }
    return rel
  })
  outputZip.file(newRelsPath, buildRelationshipsXml(updated))
}

function ensureLayoutAssets ({
  templateZip,
  outputZip,
  templateLayoutPath,
  state,
}) {
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
  const masterPath = resolveRelationshipPath(
    templateLayoutPath,
    masterRel.Target,
  )
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

function ensureMasterAssets ({
  templateZip,
  outputZip,
  templateMasterPath,
  state,
}) {
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
        const themePath = resolveRelationshipPath(
          templateMasterPath,
          rel.Target,
        )
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
        const layoutPath = resolveRelationshipPath(
          templateMasterPath,
          rel.Target,
        )
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

function copyLayoutFromMaster ({
  templateZip,
  outputZip,
  templateLayoutPath,
  newMasterPath,
  state,
}) {
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
        const mediaPath = resolveRelationshipPath(
          templateLayoutPath,
          rel.Target,
        )
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

function ensureThemeAssets ({
  templateZip,
  outputZip,
  templateThemePath,
  state,
}) {
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
  const ext = path.extname(sourcePath) || ".bin"
  const newName = `ppt/media/template_media_${state.mediaCounter++}${ext}`
  outputZip.file(newName, file.asNodeBuffer())
  state.mediaMap.set(sourcePath, newName)
  return newName
}

function copyGenericPart ({ templateZip, outputZip, sourcePath, state }) {
  if (!sourcePath.startsWith("ppt/")) {
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
  if (sourcePath.startsWith("ppt/notesSlides/")) {
    ensureContentType(
      state,
      `/${sourcePath}`,
      "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
    )
  }
  return sourcePath
}

// --- 核心修复：Content Type 检查逻辑 ---
function ensureDefaultContentType (state, extension, contentType) {
  const ext = extension.replace(/^\./, "").toLowerCase()

  // 修复：使用正则检查 Extension 属性，不限定顺序，支持单双引号
  // 匹配 <Default ... Extension="png" ...> 或 <Default ... Extension='png' ...>
  const regex = new RegExp(`<Default[^>]*Extension=["']${ext}["'][^>]*>`, "i")

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
  const partNameRegex = new RegExp(
    `<Override[^>]*PartName=["']${partName}["'][^>]*>`,
    "i",
  )
  if (state.contentTypesContent.match(partNameRegex)) {
    return
  }
  const override = `  <Override PartName="${partName}" ContentType="${contentType}"/>`
  const typesEndIndex = state.contentTypesContent.lastIndexOf("</Types>")
  if (typesEndIndex === -1) {
    throw new Error("Content_Types.xml缺少</Types>结束标签")
  }
  state.contentTypesContent =
    state.contentTypesContent.slice(0, typesEndIndex) +
    `\n${override}\n` +
    state.contentTypesContent.slice(typesEndIndex)
}

function updatePresentationDocuments (zip, state) {
  if (!state.newSlides.length) {
    return
  }
  const slideEntries = extractExistingSlideEntries(state.presContent)
  const updatedSlideEntries = buildSlideEntries(slideEntries, state)
  state.presContent = state.presContent.replace(
    slideEntries.raw,
    updatedSlideEntries.xml,
  )
  const updatedRels = appendSlideRelationships(state.presRelsContent, state)
  state.presRelsContent = updatedRels
  const updatedMasters = appendMasterEntries(state.presContent, state)
  state.presContent = updatedMasters
  for (const slide of state.newSlides) {
    ensureContentType(
      state,
      `/ppt/slides/slide${slide.slideNumber}.xml`,
      CONTENT_TYPES.slide,
    )
  }
  for (const [templatePath, layoutInfo] of state.layoutMap.entries()) {
    ensureContentType(
      state,
      `/${layoutInfo.newPath}`,
      CONTENT_TYPES.slideLayout,
    )
  }
  for (const [templatePath, masterInfo] of state.masterMap.entries()) {
    ensureContentType(
      state,
      `/${masterInfo.newPath}`,
      CONTENT_TYPES.slideMaster,
    )
  }
  for (const [templatePath, themeInfo] of state.themeMap.entries()) {
    ensureContentType(state, `/${themeInfo.newPath}`, CONTENT_TYPES.theme)
  }
  zip.file("ppt/presentation.xml", state.presContent)
  zip.file("ppt/_rels/presentation.xml.rels", state.presRelsContent)
  zip.file("[Content_Types].xml", state.contentTypesContent)
  state.newSlides = []
}

function extractExistingSlideEntries (presContent) {
  const match = presContent.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/) || {
    raw: "<p:sldIdLst></p:sldIdLst>",
    inner: "",
  }
  return { raw: match[0], inner: match[1] }
}

function buildSlideEntries (existing, state) {
  const existingEntries = existing.inner
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
  const newSlides = state.newSlides.map((slide) => {
    while (
      state.existingSlideIds.has(state.nextSlideId) ||
      state.nextSlideId <= 256
    ) {
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
  const entryFor = (slide) =>
    `    <p:sldId id="${slide.slideId}" r:id="${slide.relId}"/>`
  const startEntries = newSlides
    .filter((s) => s.position === "start")
    .map(entryFor)
  const middleEntries = newSlides
    .filter((s) => s.position === "middle")
    .map(entryFor)
  const endEntries = newSlides
    .filter((s) => s.position === "end")
    .map(entryFor)
  // 正确顺序：start -> existing -> middle -> end
  const xml = `<p:sldIdLst>
${[...startEntries, ...existingEntries, ...middleEntries, ...endEntries].join("\n")}
</p:sldIdLst>`
  return { xml }
}

function appendSlideRelationships (relsContent, state) {
  const insertIndex = relsContent.indexOf("</Relationships>")
  if (insertIndex === -1) {
    throw new Error("presentation.xml.rels 内容异常")
  }
  const slideRels = state.newSlides
    .map(
      (slide) =>
        `  <Relationship Id="${slide.relId}" Type="${REL_TYPES.slide}" Target="slides/slide${slide.slideNumber}.xml"/>`,
    )
    .join("\n")
  return `${relsContent.slice(0, insertIndex)}${slideRels}\n${relsContent.slice(insertIndex)}`
}

function appendMasterEntries (presContent, state) {
  if (!state.newMastersList.length) {
    return presContent
  }
  const relsContent = state.presRelsContent
  const insertIndex = relsContent.indexOf("</Relationships>")
  if (insertIndex === -1) {
    throw new Error("presentation.xml.rels 缺少 Relationships 结束标签")
  }
  const masterListMatch = presContent.match(
    /<p:sldMasterIdLst>([\s\S]*?)<\/p:sldMasterIdLst>/,
  ) || { 1: "" }
  const existingEntries = masterListMatch[1]
    ? masterListMatch[1]
      .split("\n")
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
        "ppt/",
        "",
      )}"/>`,
    )
    state.existingMasterIds.add(state.nextMasterId)
    state.existingRelIds.add(state.nextRelId)
    state.nextMasterId++
    state.nextRelId++
  }
  const newMasterList = `<p:sldMasterIdLst>
${[...existingEntries, ...newMasterEntries].join("\n")}
</p:sldMasterIdLst>`
  let updatedPres = presContent
  if (masterListMatch[0]) {
    updatedPres = presContent.replace(masterListMatch[0], newMasterList)
  } else {
    updatedPres = presContent.replace(
      "</p:presentation>",
      `${newMasterList}\n</p:presentation>`,
    )
  }
  const updatedRelsContent = `${relsContent.slice(
    0,
    insertIndex,
  )}${newMasterRels.join("\n")}\n${relsContent.slice(insertIndex)}`
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
        .join(" ")
      return `  <Relationship ${attrs}/>`
    })
    .join("\n")
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
  relPath = relPath.replace(/\\/g, "/")
  relPath = relPath.replace(/^\.\//g, "")
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

function applyCoverPlaceholders (xml, employee = {}) {
  const replacements = {
    姓名: employee.name || "",
    性别: employee.gender || "",
    年龄: employee.age || "",
    日期: employee.date || "",
    工号: employee.id || "",
  }
  return replaceTextInSlidePreferred(xml, replacements)
}

function formatCnDate (dateInput) {
  const date =
    dateInput instanceof Date ? dateInput : new Date(dateInput || Date.now())
  if (Number.isNaN(date.getTime())) {
    return ""
  }
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, "0")
  const day = String(date.getDate()).padStart(2, "0")
  return `${year}年${month}月${day}日`
}

function escapeXmlValue (value) {
  if (!value) return ""
  return String(value)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;")
}

function replaceTextInSlidePreferred (xmlContent, replacements) {
  let result = xmlContent
  const optionalSpaces = "(?:\\s|\\u00A0|&nbsp;|&#160;|&#xA0;)*"
  for (const [key, value] of Object.entries(replacements)) {
    if (value === undefined || value === null) continue
    const valueStr = String(value)
    if (!valueStr && key !== "性别" && key !== "年龄") continue
    const escapedValue = escapeXmlValue(valueStr)
    const keyPattern = escapeRegex(key)
    const textRunPattern = /(<a:t[^>]*>)([\s\S]*?)(<\/a:t>)/gi
    result = result.replace(
      textRunPattern,
      (match, openTag, content, closeTag) => {
        let updatedContent = content
        const placeholderPattern = new RegExp(
          `\\{\\{${optionalSpaces}${keyPattern}${optionalSpaces}\\}\}${optionalSpaces}`,
          "gi",
        )
        updatedContent = updatedContent.replace(
          placeholderPattern,
          (phMatch) => {
            const suffixMatch = phMatch.match(
              /((?:\\s|\\u00A0|&nbsp;|&#160;|&#xA0;)*)$/i,
            )
            const suffix = suffixMatch ? suffixMatch[1] : ""
            return `${escapedValue}${suffix}`
          },
        )
        const bracketPattern = new RegExp(
          `\\[${optionalSpaces}${keyPattern}${optionalSpaces}\\]`,
          "gi",
        )
        updatedContent = updatedContent.replace(bracketPattern, escapedValue)
        const directPattern = new RegExp(
          `${optionalSpaces}${keyPattern}${optionalSpaces}`,
          "gi",
        )
        updatedContent = updatedContent.replace(directPattern, escapedValue)
        return `${openTag}${updatedContent}${closeTag}`
      },
    )
  }
  return result
}

function escapeRegex (str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
}

// --- 并发控制工具函数 ---
async function limitConcurrency (tasks, maxConcurrency) {
  const results = []
  const executing = []
  for (const task of tasks) {
    const promise = Promise.resolve(task()).then((result) => {
      executing.splice(executing.indexOf(promise), 1)
      return result
    })
    results.push(promise)
    executing.push(promise)
    if (executing.length >= maxConcurrency) {
      await Promise.race(executing)
    }
  }
  return Promise.all(results)
}

// --- PDF 转换并发控制队列 ---
class PdfConversionQueue {
  constructor(maxConcurrency) {
    this.maxConcurrency = maxConcurrency
    this.running = 0
    this.queue = []
  }

  async add (task) {
    return new Promise((resolve, reject) => {
      this.queue.push({ task, resolve, reject })
      this.process()
    })
  }

  async process () {
    // 如果已达到最大并发数或队列为空，直接返回
    if (this.running >= this.maxConcurrency || this.queue.length === 0) {
      return
    }

    // 取出一个任务
    const { task, resolve, reject } = this.queue.shift()
    this.running++

    try {
      const result = await task()
      resolve(result)
    } catch (error) {
      reject(error)
    } finally {
      this.running--
      // 处理下一个任务
      setImmediate(() => this.process())
    }
  }
}

// 创建全局 PDF 转换队列实例
const pdfConversionQueue = new PdfConversionQueue(
  MAX_CONCURRENT_PDF_CONVERSIONS,
)

// 模板文件缓存
let templateBufferCache = null
let templateZipCache = null

async function getTemplateBuffer (templatePath) {
  if (!templateBufferCache) {
    templateBufferCache = await fs.readFile(templatePath)
    templateZipCache = new PizZip(templateBufferCache)
  }
  return { buffer: templateBufferCache, zip: templateZipCache }
}

function clearTemplateCache () {
  templateBufferCache = null
  templateZipCache = null
}

function replaceImagePlaceholder (
  slideXml,
  image,
  state,
  outputZip,
  slideNumber,
) {
  try {
    let ext = ".png"
    let mimeType = "image/png"
    if (image.data) {
      const dataUrlMatch = image.data.match(/^data:(image\/\w+);base64,/)
      if (dataUrlMatch) {
        mimeType = dataUrlMatch[1]
        ext = `.${mimeType.split("/")[1]}`
      }
    } else if (image.fullPath) {
      ext = path.extname(image.fullPath) || ".png"
      mimeType = ext.toLowerCase().includes("png")
        ? "image/png"
        : ext.toLowerCase().includes("jpg") ||
          ext.toLowerCase().includes("jpeg")
          ? "image/jpeg"
          : ext.toLowerCase().includes("gif")
            ? "image/gif"
            : ext.toLowerCase().includes("bmp")
              ? "image/bmp"
              : "image/png"
    } else {
      console.warn(`⚠️ 图片数据为空，跳过处理`)
      return slideXml
    }
    const newImagePath = `ppt/media/image_${state.mediaCounter++}${ext}`
    ensureDefaultContentType(state, ext, mimeType)
    let buffer
    if (image.data) {
      const base64Data = image.data.replace(/^data:image\/\w+;base64,/, "")
      buffer = Buffer.from(base64Data, "base64")
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
    updatedXml = updatedXml.replace(/\{\{\s*图片\s*\}\}/g, "")
    return updatedXml
  } catch (err) {
    console.warn(`⚠️ 替换幻灯片图片失败：${err.message}`)
    return slideXml
  }
}

// 按照模板顺序插入图片页面：inbody -> lab -> ecg -> ai
async function insertImageSlidesByCategory (
  templatePath,
  outputPath,
  imageItems,
  employee,
  category,
  templateSlideNumber,
) {
  const categoryImages = imageItems.filter((img) => img.category === category)
  if (categoryImages.length === 0) {
    return
  }

  const [templateData, outputBuffer] = await Promise.all([
    getTemplateBuffer(templatePath),
    fs.readFile(outputPath),
  ])
  const templateZip = templateData.zip
  const outputZip = new PizZip(outputBuffer)
  const state = initializeState(outputZip, { employee, date: new Date() })
  const newSlides = []

  for (const image of categoryImages) {
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
      姓名: employee.name || "",
      工号: employee.id || "",
    }
    slideXml = replaceTextInSlidePreferred(slideXml, replacements)
    slideXml = replaceImagePlaceholder(
      slideXml,
      image,
      state,
      outputZip,
      newSlideNumber,
    )
    outputZip.file(newSlidePath, slideXml)
    cloneSlideRelationships({
      templateZip,
      outputZip,
      templateSlideNumber,
      newSlideNumber,
      state,
    })
    ensureContentType(
      state,
      `/ppt/slides/slide${newSlideNumber}.xml`,
      CONTENT_TYPES.slide,
    )
    newSlides.push({
      slideNumber: newSlideNumber,
      position: "middle", // 图片页面应该在封面之后，总结页之前
    })
  }

  if (!newSlides.length) {
    return
  }
  state.newSlides = newSlides
  updatePresentationDocuments(outputZip, state)
  const updatedBuffer = outputZip.generate({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  })
  await fs.writeFile(outputPath, updatedBuffer)
}

// --- PPTX 转 PDF 核心转换函数（内部使用，带重试机制）---
async function _convertPptxToPdfInternal (pptxPath, retryCount = 0) {
  const MAX_RETRIES = 2
  const RETRY_DELAY = 1000 // 重试延迟1秒

  const pdfPath = pptxPath.replace(/\.pptx$/i, ".pdf")

  // 添加小延迟，避免LibreOffice进程冲突
  if (retryCount === 0) {
    await new Promise((resolve) => setTimeout(resolve, 200))
  }

  const pptxBuf = await fs.readFile(pptxPath)
  try {
    const pdfBuf = await convertAsync(pptxBuf, ".pdf", "impress_pdf_Export")
    await fs.writeFile(pdfPath, pdfBuf)
  } catch (err) {
    // 如果 libreoffice-convert 失败，使用 soffice 命令行
    const candidates = [
      process.env.LIBREOFFICE_PATH,
      process.env.SOFFICE_PATH,
      "C:/Program Files/LibreOffice/program/soffice.exe",
      "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
    ].filter(Boolean)
    let sofficePath = null
    for (const p of candidates) {
      try {
        if (p && (await fs.pathExists(p))) {
          sofficePath = p
          break
        }
      } catch { }
    }
    if (!sofficePath) {
      throw new Error(
        "未找到 soffice，可设置 LIBREOFFICE_PATH 或 SOFFICE_PATH 指向 soffice.exe",
      )
    }

    const outDir = path.dirname(pdfPath)
    await fs.ensureDir(outDir)

    // 清理可能存在的旧PDF文件
    if (await fs.pathExists(pdfPath)) {
      try {
        await fs.remove(pdfPath)
      } catch { }
    }

    await new Promise((resolve, reject) => {
      const args = [
        "--headless",
        "--invisible",
        "--nodefault",
        "--nolockcheck",
        "--convert-to",
        'pdf:impress_pdf_Export:{"ReduceImageResolution":false,"Quality":100}',
        "--outdir",
        outDir,
        pptxPath,
      ]
      const child = spawn(sofficePath, args, {
        stdio: ["ignore", "pipe", "pipe"],
        windowsHide: true,
      })
      let stderr = ""
      let stdout = ""
      child.stdout.on("data", (d) => {
        stdout += String(d)
      })
      child.stderr.on("data", (d) => {
        stderr += String(d)
      })
      child.on("error", (e) => reject(e))
      child.on("close", (code) => {
        if (code === 0) {
          resolve()
        } else {
          const errorMsg = stderr.trim() || stdout.trim() || "未知错误"
          reject(new Error(`soffice 退出码 ${code}: ${errorMsg}`))
        }
      })
    })

    // 等待文件生成
    let attempts = 0
    while (attempts < 10 && !(await fs.pathExists(pdfPath))) {
      await new Promise((resolve) => setTimeout(resolve, 200))
      attempts++
    }

    if (!(await fs.pathExists(pdfPath))) {
      if (retryCount < MAX_RETRIES) {
        console.warn(
          `⚠️ PDF 文件未生成，${RETRY_DELAY}ms 后重试 (${retryCount + 1}/${MAX_RETRIES})...`,
        )
        await new Promise((resolve) => setTimeout(resolve, RETRY_DELAY))
        return _convertPptxToPdfInternal(pptxPath, retryCount + 1)
      }
      throw new Error("soffice 执行后未生成 PDF 文件")
    }

    const stat = await fs.stat(pdfPath)
    if (!stat.size) {
      if (retryCount < MAX_RETRIES) {
        console.warn(
          `⚠️ PDF 文件为空，${RETRY_DELAY}ms 后重试 (${retryCount + 1}/${MAX_RETRIES})...`,
        )
        await new Promise((resolve) => setTimeout(resolve, RETRY_DELAY))
        return _convertPptxToPdfInternal(pptxPath, retryCount + 1)
      }
      throw new Error("生成的 PDF 为空")
    }
  }
  return pdfPath
}

// --- PPTX 转 PDF 公共接口（带并发控制）---
async function convertPptxToPdf (pptxPath) {
  return await pdfConversionQueue.add(() =>
    _convertPptxToPdfInternal(pptxPath),
  )
}
// --------------------------------

// 新增：转换result_add_suggest文件夹内的PPT为PDF
async function convertResultPptToPdf () {
  const RESULT_DIR = path.join(ROOT, "result_add_suggest")
  if (!(await fs.pathExists(RESULT_DIR))) {
    console.error(`❌ result_add_suggest 文件夹不存在`)
    return
  }

  const files = await fs.readdir(RESULT_DIR)
  const pptxFiles = files.filter(
    (file) => path.extname(file).toLowerCase() === ".pptx",
  )

  if (pptxFiles.length === 0) {
    console.log(`ℹ️ result_add_suggest 文件夹内没有PPT文件`)
    return
  }

  console.log(`📋 找到 ${pptxFiles.length} 个PPT文件，开始转换...`)

  let successCount = 0
  let failCount = 0

  for (const pptxFile of pptxFiles) {
    const pptxPath = path.join(RESULT_DIR, pptxFile)
    try {
      console.log(`⏳ 正在转换: ${pptxFile}...`)
      const pdfPath = await convertPptxToPdf(pptxPath)
      console.log(`✓ PDF 生成成功: ${path.basename(pdfPath)}`)
      successCount++
    } catch (pdfErr) {
      console.error(`❌ PDF 转换失败 (${pptxFile}): ${pdfErr.message}`)
      failCount++
    }
  }

  console.log(`\n===== 转换统计 =====`)
  console.log(`✅ 成功: ${successCount} 个`)
  console.log(`❌ 失败: ${failCount} 个`)
}

async function processEmployee (
  employee,
  availableFiles,
  templatePath,
  layout,
  nameCounter,
) {
  const assetInfo = await collectEmployeeAssets(employee, availableFiles)
  const hasSummary = Boolean(assetInfo.summaryText.trim())
  const hasAttachments = assetInfo.attachments.length > 0
  if (!hasSummary && !hasAttachments) {
    return {
      success: false,
      employee,
      reason: "缺少体检结果与AI总结",
    }
  }

  const pptx = initializePresentation(layout)
  const currentCount = nameCounter.get(employee.name) || 0
  const suffix = currentCount > 0 ? `_${currentCount}` : ""
  const outputName = buildReportFileName(employee, suffix)
  const outputPath = path.join(OUTPUT_DIR, outputName)
  nameCounter.set(employee.name, currentCount + 1)

  await pptx.writeFile({ fileName: outputPath })

  const options = {
    employee,
    date: new Date(),
    summary: assetInfo.summaryText,
  }

  // 1. 先插入模板第1页（封面）
  await insertCoverSlide(templatePath, outputPath, options)

  // 2. 按照模板顺序插入图片页面：inbody -> lab -> ecg -> ai
  const imageItems = await buildImageItems(assetInfo, employee)
  if (imageItems.length > 0) {
    // 按照模板顺序逐个插入
    await insertImageSlidesByCategory(
      templatePath,
      outputPath,
      imageItems,
      employee,
      "inbody",
      2,
    )
    await insertImageSlidesByCategory(
      templatePath,
      outputPath,
      imageItems,
      employee,
      "lab",
      3,
    )
    await insertImageSlidesByCategory(
      templatePath,
      outputPath,
      imageItems,
      employee,
      "ecg",
      4,
    )
    await insertImageSlidesByCategory(
      templatePath,
      outputPath,
      imageItems,
      employee,
      "ai",
      5,
    )
  }

  // 3. 插入建议与指导（文字）第6页（如果有 has6WhitePage）
  if (has6WhitePage) {
    await insertTextGuidanceSlide(templatePath, outputPath, options)
  }

  // 4. 最后插入结束页（第7页）
  await insertSummarySlides(templatePath, outputPath, options)

  // --- PDF 转换逻辑 ---
  let finalOutputPath = outputPath
  if (generatePdf) {
    try {
      console.log(`⏳ 正在转换为 PDF: ${outputName}...`)
      const pdfPath = await convertPptxToPdf(outputPath)
      console.log(`✓ PDF 生成成功: ${path.basename(pdfPath)}`)
      try {
        const stat = await fs.stat(pdfPath)
        if (stat.size > 0) {
          await fs.remove(outputPath)
          console.log(`✓ 已删除 PPTX: ${path.basename(outputPath)}`)
          // 更新最终输出路径为PDF路径
          finalOutputPath = pdfPath
        }
      } catch (delErr) {
        console.warn(`⚠️ 删除 PPTX 失败: ${delErr.message}`)
      }
    } catch (pdfErr) {
      console.error(`❌ PDF 转换失败 (${employee.name}): ${pdfErr.message}`)
      console.error(
        `   请确保系统已安装 LibreOffice 并且 npm install libreoffice-convert 已运行。`,
      )
      // PDF转换失败时，保留PPTX路径
    }
  }
  // ------------------------

  return {
    success: true,
    employee,
    outputPath: finalOutputPath,
  }
}

async function main () {
  const startTime = Date.now()
  await fs.ensureDir(OUTPUT_DIR)
  await fs.ensureDir(PDF_IMAGE_DIR)
  await fs.ensureDir(PDF_TEMP_DIR)
  const templatePath = await resolveExistingPath(
    TEMPLATE_CANDIDATES,
    "模板文件",
  )
  const sheetPath = await resolveExistingPath(
    EMPLOYEE_SHEET_CANDIDATES,
    "员工表",
  )
  const layout = await loadTemplateLayout(templatePath)
  const employees = await loadEmployees(sheetPath)
  if (!employees.length) {
    console.warn("⚠️ 员工表为空，已结束。")
    return
  }

  // 预加载模板文件到缓存
  await getTemplateBuffer(templatePath)

  const availableFiles = await safeReadDir(DATA_DIR)
  const successReports = []
  const skippedEmployees = []
  const nameCounter = new Map()

  // 创建处理任务
  const tasks = employees.map((employee) => async () => {
    try {
      const result = await processEmployee(
        employee,
        availableFiles,
        templatePath,
        layout,
        nameCounter,
      )
      if (result.success) {
        successReports.push({
          employee: result.employee,
          outputPath: result.outputPath,
        })
        console.log(
          `✓ 已生成 ${result.employee.name}（${result.employee.id}）：${result.outputPath}`,
        )
      } else {
        skippedEmployees.push({
          employee: result.employee,
          reason: result.reason,
        })
      }
      return result
    } catch (error) {
      skippedEmployees.push({
        employee,
        reason: `生成失败：${error.message}`,
      })
      console.error(`❌ ${employee.name} 生成失败：${error.message}`)
      return { success: false, employee, error }
    }
  })

  // 使用并发控制处理所有员工
  console.log(
    `🚀 开始处理 ${employees.length} 个员工，并发数：${MAX_CONCURRENT_EMPLOYEES}`,
  )
  await limitConcurrency(tasks, MAX_CONCURRENT_EMPLOYEES)

  // 清理模板缓存
  clearTemplateCache()

  const endTime = Date.now()
  const duration = ((endTime - startTime) / 1000).toFixed(2)

  console.log("\n===== 生成统计 =====")
  console.log(`✅ 已生成：${successReports.length} 人`)
  successReports.forEach((item) => {
    console.log(
      `  - ${item.employee.name}（${item.employee.id}） -> ${item.outputPath}`,
    )
  })
  console.log(`⚠️ 未生成：${skippedEmployees.length} 人`)
  skippedEmployees.forEach((item) => {
    console.log(
      `  - ${item.employee.name}（${item.employee.id}）：${item.reason}`,
    )
  })
  console.log(`⏱️ 总耗时：${duration} 秒`)
}

// 新增：处理result_add_suggest文件夹下的PPT文件名，添加证件号
async function renameResultPptFiles () {
  const RESULT_DIR = path.join(ROOT, "result_add_suggest")
  if (!(await fs.pathExists(RESULT_DIR))) {
    console.error(`❌ result_add_suggest 文件夹不存在`)
    return
  }

  // 读取员工表数据
  const sheetPath = await resolveExistingPath(
    EMPLOYEE_SHEET_CANDIDATES,
    "员工表",
  )
  const employees = await loadEmployees(sheetPath)

  // 建立姓名到证件号的映射
  const nameToIdMap = new Map()
  for (const employee of employees) {
    const currentIds = nameToIdMap.get(employee.name) || []
    currentIds.push(employee.id)
    nameToIdMap.set(employee.name, currentIds)
  }

  // 获取文件夹中的PPT文件
  const files = await fs.readdir(RESULT_DIR)
  const pptxFiles = files.filter(
    (file) => path.extname(file).toLowerCase() === ".pptx",
  )

  if (pptxFiles.length === 0) {
    console.log(`ℹ️ result_add_suggest 文件夹内没有PPT文件`)
    return
  }

  console.log(`📋 找到 ${pptxFiles.length} 个PPT文件，开始重命名...`)

  let successCount = 0
  let skipCount = 0

  for (const pptxFile of pptxFiles) {
    try {
      // 从文件名中提取姓名
      const nameMatch = pptxFile.match(
        /^体检报告_([^_]+?)(?:_\d+)?(?:_[\dXx]+)?\.pptx$/i,
      )
      if (!nameMatch) {
        console.warn(`⚠️ 文件名格式不符合要求：${pptxFile}，跳过`)
        skipCount++
        continue
      }

      let name = nameMatch[1]
      name = normalizeText(name) // 规范化姓名，确保与员工表中的姓名匹配
      console.log(`⏳ 处理文件：${pptxFile}`)
      console.log(`   提取姓名：${name}`)

      // 查询证件号
      const ids = nameToIdMap.get(name)
      if (!ids || ids.length === 0) {
        console.warn(`⚠️ 未找到姓名为 "${name}" 的员工，跳过`)
        skipCount++
        continue
      }

      if (ids.length > 1) {
        console.warn(`⚠️ 姓名为 "${name}" 的员工有多个证件号，跳过`)
        skipCount++
        continue
      }

      const id = ids[0]
      console.log(`   匹配证件号：${id}`)

      // 检查文件名是否已经包含证件号
      if (pptxFile.includes(`_${id}`)) {
        console.warn(`⚠️ 文件 ${pptxFile} 已经包含证件号，跳过`)
        skipCount++
        continue
      }

      // 生成新文件名
      const baseName = path.basename(pptxFile, ".pptx")
      const newFileName = `${baseName}_${id}.pptx`
      const oldPath = path.join(RESULT_DIR, pptxFile)
      const newPath = path.join(RESULT_DIR, newFileName)

      // 执行重命名
      await fs.rename(oldPath, newPath)
      console.log(`✅ 重命名成功：${pptxFile} -> ${newFileName}`)
      successCount++
    } catch (error) {
      console.error(`❌ 重命名失败 (${pptxFile}): ${error.message}`)
      skipCount++
    }
  }

  console.log(`\n===== 重命名统计 =====`)
  console.log(`✅ 成功：${successCount} 个`)
  console.log(`⚠️ 跳过：${skipCount} 个`)
}

// 将文件通过身份证号匹配然后发送企业微信
async function sendWechatEeMsg () {
  const SEND_DATA_DIR = path.join(ROOT, "send_data")
  if (!(await fs.pathExists(SEND_DATA_DIR))) {
    console.error(`❌ send_data 文件夹不存在`)
    return
  }

  const files = await fs.readdir(SEND_DATA_DIR)
  const pdfFiles = files.filter(
    (file) => path.extname(file).toLowerCase() === ".pdf",
  )

  if (pdfFiles.length === 0) {
    console.log(`ℹ️ send_data 文件夹内没有PDF文件`)
    return
  }

  console.log(`📋 找到 ${pdfFiles.length} 个PDF文件，开始提取证件号...`)

  let successCount = 0
  let failCount = 0

  // 导入axios和form-data库
  const axios = require("axios")
  const FormData = require("form-data")

  // 定义发送POST请求获取企信ID的函数
  async function getQixinId (sfz) {
    try {
      // 使用axios发送POST请求，将sfz参数放在URL中
      const response = await axios.post(
        `http://wxsite.yinda.cn:5182/cajserver/pro/caj-renlizy/ZhiGong/nologin/getQixinIdBySfz?sfz=${encodeURIComponent(sfz)}`,
        {},
        {
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
          },
        },
      )
      return response.data
    } catch (error) {
      if (error.response) {
        throw new Error(
          `请求失败: ${error.response.status} ${error.response.statusText}`,
        )
      } else if (error.request) {
        throw new Error(`请求失败: 没有收到响应`)
      } else {
        throw new Error(`请求失败: ${error.message}`)
      }
    }
  }

  // 定义发送form-data请求上传文件的函数
  async function sendFileToUser (userId, fileFullPath, fileName) {
    try {
      // 创建form-data对象
      const formData = new FormData()
      // 添加文件
      formData.append("file", fs.createReadStream(fileFullPath), {
        filename: fileName,
        contentType: "application/pdf",
      })

      // 使用axios发送POST请求
      const response = await axios.post(
        `https://product.cajcare.com:5182/wechat/caj/sunflower/sendFileToUser?userId=${encodeURIComponent(userId)}`,
        formData,
        {
          headers: {
            ...formData.getHeaders(),
          },
        },
      )
      return response.data
    } catch (error) {
      if (error.response) {
        throw new Error(
          `发送文件失败: ${error.response.status} ${error.response.statusText}`,
        )
      } else if (error.request) {
        throw new Error(`发送文件失败: 没有收到响应`)
      } else {
        throw new Error(`发送文件失败: ${error.message}`)
      }
    }
  }

  for (const pdfFile of pdfFiles) {
    try {
      // 从文件名中提取证件号，命名规则：体检报告_姓名_证件号.pdf 或 体检报告_姓名_证件号_数字.pdf
      const idMatch = pdfFile.match(
        /^体检报告_([^_]+)_([\dXx]+)(?:_\d+)?\.pdf$/,
      )
      if (!idMatch) {
        console.warn(`⚠️ 文件名格式不符合要求：${pdfFile}`)
        failCount++
        continue
      }

      const name = idMatch[1]
      const id = idMatch[2]
      const fileFullPath = path.join(SEND_DATA_DIR, pdfFile)
      console.log(`✅ ${pdfFile} -> 姓名：${name}，证件号：${id}`)

      // 调用接口获取企信ID
      console.log(`⏳ 正在获取${name}的企信ID...`)
      const qixinResult = await getQixinId(id)

      if (qixinResult.returnCode === 1) {
        const qixinId = qixinResult.returnData
        console.log(`✅ 企信ID获取成功：${qixinId}`)

        // 调用接口发送PDF文件
        console.log(`⏳ 正在发送${pdfFile}到企信...`)
        const sendResult = await sendFileToUser(qixinId, fileFullPath, pdfFile)

        if (sendResult.code === 1) {
          console.log(`✅ 文件发送成功：${sendResult.message}`)
        } else {
          console.warn(`⚠️ 文件发送失败：${sendResult.message || "未知错误"}`)
        }
      } else {
        console.warn(`⚠️ 企信ID获取失败：${qixinResult.returnMessage}`)
      }

      successCount++
    } catch (error) {
      console.error(`❌ 处理失败 (${pdfFile}): ${error.message}`)
      failCount++
    }
  }

  console.log(`\n===== 提取统计 =====`)
  console.log(`✅ 成功：${successCount} 个`)
  console.log(`❌ 失败：${failCount} 个`)
}

// 检查命令行参数
const args = process.argv.slice(2)
if (args.includes("--convert-ppt-pdf")) {
  // 执行PPT转PDF命令
  convertResultPptToPdf().catch((error) => {
    console.error("❌ 转换失败：", error)
    process.exitCode = 1
  })
} else if (args.includes("--rename-ppt-files")) {
  // 执行PPT文件名重命名命令
  renameResultPptFiles().catch((error) => {
    console.error("❌ 重命名失败：", error)
    process.exitCode = 1
  })
} else {
  // 执行原始主流程
  main().catch((error) => {
    console.error("❌ 生成失败：", error)
    process.exitCode = 1
  })
}
