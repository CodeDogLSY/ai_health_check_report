const path = require('path')
const fs = require('fs-extra')
const crypto = require('crypto')
const mammoth = require('mammoth')
const { fromPath: pdfFromPath } = require('pdf2pic')
const { createCanvas } = require('canvas')
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js')

const {
  DATA_DIR,
  PDF_TEMP_DIR,
  PDF_IMAGE_DIR,
  IMAGE_EXTENSIONS,
  PDF_EXTENSIONS,
} = require('./config')
const {
  normalizeName,
  buildAsciiSafeLabel,
} = require('./utils')

pdfjsLib.GlobalWorkerOptions.workerSrc = require.resolve('pdfjs-dist/legacy/build/pdf.worker.js')

/**
 * 读取数据目录，若不存在则返回空数组。
 */
async function safeReadDir (dir) {
  if (!(await fs.pathExists(dir))) {
    console.warn(`⚠️ 数据目录不存在：${dir}`)
    return []
  }
  return fs.readdir(dir)
}

/**
 * 为单个员工收集所有匹配的资料文件。
 */
async function collectEmployeeAssets (employee, files) {
  const normalized = normalizeName(employee.name)
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

      // 确定文件优先级，按照 inbody、尿检、血检、心电图、AI解读的顺序
      let priority = 99 // 默认优先级
      if (file.toLowerCase().includes('inbody')) {
        priority = 1
      } else if (file.toLowerCase().includes('尿常规') || file.toLowerCase().includes('尿检')) {
        priority = 2
      } else if (file.toLowerCase().includes('血检')) {
        priority = 3
      } else if (file.toLowerCase().includes('心电图')) {
        priority = 4
      } else if (file.toLowerCase().includes('ai解读') || file.toLowerCase().includes('ai解读')) {
        priority = 5
      }

      return {
        fileName: file,
        fullPath: path.join(DATA_DIR, file),
        label,
        type,
        ext,
        priority,
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

/**
 * 将图片和 PDF（转换为图片）统筹为 PPT 可用素材。
 */
async function buildImageItems (assetInfo, employee) {
  const imageItems = []

  // 按照附件的优先级顺序处理每个附件
  for (const attachment of assetInfo.attachments) {
    if (attachment.type === 'image') {
      // 图片类型直接添加
      imageItems.push({
        label: attachment.label,
        fullPath: attachment.fullPath,
      })
    } else if (attachment.type === 'pdf') {
      // PDF类型转换为图片后添加
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

    const pdfImgPages = await convertWithPdfRenderer(tempPdfPath, pdfAttachment)
    if (pdfImgPages.length) {
      return pdfImgPages
    }

    return convertWithPdf2Pic(tempPdfPath, pdfAttachment, safeLabel)
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

module.exports = {
  safeReadDir,
  collectEmployeeAssets,
  buildImageItems,
};

