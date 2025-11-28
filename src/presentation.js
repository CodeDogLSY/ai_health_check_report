const PptxGenJS = require('pptxgenjs')
const { sanitizeForFilename, formatDate } = require('./utils')

/**
 * 初始化 PPT 文档及基础元数据。
 */
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

function getLayoutGuides (layout) {
  const sideMargin = Number((layout.width * 0.08).toFixed(3))
  const topMargin = Number((layout.height * 0.05).toFixed(3))
  const bottomMargin = Number((layout.height * 0.06).toFixed(3))
  const contentWidth = Number((layout.width - sideMargin * 2).toFixed(3))
  return { sideMargin, topMargin, bottomMargin, contentWidth }
}

function addOverviewSlide (pptx, employee, assets, theme, layout) {
  const slide = pptx.addSlide()
  slide.background = { color: 'FFFFFF' }

  const metrics = getLayoutGuides(layout)
  slide.addText(`${employee.name} · 检查项概览`, {
    x: metrics.sideMargin,
    y: metrics.topMargin,
    w: metrics.contentWidth,
    fontSize: layout.orientation === 'portrait' ? 28 : 24,
    bold: true,
    color: theme.primary.replace('#', ''),
  })

  const headerCells = [
    { text: '类别', options: { bold: true, color: theme.textLight.replace('#', ''), fill: theme.primary.replace('#', '') } },
    { text: '文件名', options: { bold: true, color: theme.textLight.replace('#', ''), fill: theme.primary.replace('#', '') } },
    { text: '备注', options: { bold: true, color: theme.textLight.replace('#', ''), fill: theme.primary.replace('#', '') } },
  ]

  const rows = [headerCells]

  if (!assets.attachments.length) {
    rows.push([
      { text: '暂无', options: {} },
      { text: '未检测到检查文件', options: {} },
      { text: '', options: {} },
    ])
  } else {
    assets.attachments.forEach((item) => {
      rows.push([
        { text: formatAttachmentType(item.type), options: {} },
        { text: item.fileName, options: {} },
        { text: item.label || '附件', options: {} },
      ])
    })
  }

  const tableY = metrics.topMargin + 0.8
  slide.addTable(rows, {
    x: metrics.sideMargin,
    y: tableY,
    w: metrics.contentWidth,
    colW: [
      Number((metrics.contentWidth * 0.18).toFixed(3)),
      Number((metrics.contentWidth * 0.5).toFixed(3)),
      Number((metrics.contentWidth * 0.32).toFixed(3)),
    ],
    border: { type: 'solid', color: theme.neutral.replace('#', ''), pt: 1 },
    fill: theme.background.replace('#', ''),
    autoPage: true,
    autoPageRepeatHeader: true,
    autoPageLineWeight: 0.5,
    color: theme.textDark.replace('#', ''),
    fontSize: 16,
    bold: false,
    margin: 0.1,
    rowH: 0.4,
    valign: 'middle',
    align: 'left',
    header: true,
    headerFill: theme.primary.replace('#', ''),
  })
}

/**
 * 添加体检总结幻灯片
 * @param {Object} pptx - PptxGenJS 实例
 * @param {Object} employee - 员工信息
 * @param {Object} assets - 员工资产信息
 * @param {Object} theme - 主题颜色
 * @param {Object} layout - 布局信息
 */
function addSummarySlide (pptx, employee, assets, theme, layout) {
  const slide = pptx.addSlide()
  slide.background = { color: theme.background.replace('#', '') }

  const metrics = getLayoutGuides(layout)
  slide.addText(`${employee.name} · 体检总结`, {
    x: metrics.sideMargin,
    y: metrics.topMargin,
    w: metrics.contentWidth,
    fontSize: layout.orientation === 'portrait' ? 28 : 24,
    bold: true,
    color: theme.primary.replace('#', ''),
  })

  const summary = assets.summaryText || '暂无总结，请补充。'
  const paragraphs = chunkText(summary, 120)
  const textRuns = paragraphs.map((paragraph) => ({
    text: paragraph,
    options: {
      bullet: true,
      color: theme.textDark.replace('#', ''),
      fontSize: 18,
    },
  }))

  slide.addText(textRuns, {
    x: metrics.sideMargin,
    y: metrics.topMargin + 0.8,
    w: metrics.contentWidth,
    h: layout.height - metrics.topMargin - metrics.bottomMargin - 1,
    lineSpacingMultiple: 1.2,
  })
}

/**
 * 添加影像资料幻灯片
 * @param {Object} pptx - PptxGenJS 实例
 * @param {Object} employee - 员工信息
 * @param {Array} imageItems - 影像资料列表
 * @param {Object} theme - 主题颜色
 * @param {Object} layout - 布局信息
 */
function addImageSlides (pptx, employee, imageItems, theme, layout) {
  if (!imageItems.length) {
    return
  }

  const metrics = getLayoutGuides(layout)
  imageItems.forEach((image, index) => {
    const slide = pptx.addSlide()
    slide.background = { color: theme.background.replace('#', '') }

    slide.addText(`${employee.name} · 影像资料 ${index + 1}/${imageItems.length}`, {
      x: metrics.sideMargin,
      y: metrics.topMargin,
      w: metrics.contentWidth,
      fontSize: layout.orientation === 'portrait' ? 26 : 22,
      bold: true,
      color: theme.primary.replace('#', ''),
    })

    slide.addText(image.label, {
      x: metrics.sideMargin,
      y: metrics.topMargin + 0.6,
      w: metrics.contentWidth,
      fontSize: layout.orientation === 'portrait' ? 20 : 18,
      color: theme.secondary.replace('#', ''),
    })

    const imagePayload = image.data
      ? { data: image.data }
      : { path: image.fullPath }

    const imageTop = metrics.topMargin + 1.1
    const maxWidth = metrics.contentWidth
    const maxHeight = Math.max(3, layout.height - imageTop - metrics.bottomMargin)

    slide.addImage({
      ...imagePayload,
      x: metrics.sideMargin,
      y: imageTop,
      w: maxWidth,
      h: maxHeight,
      sizing: { type: 'contain', w: maxWidth, h: maxHeight },
    })
  })
}

function addDocumentSlide (pptx, employee, attachments, theme, layout) {
  const documents = attachments.filter((item) => item.type !== 'image')
  if (!documents.length) {
    return
  }

  const slide = pptx.addSlide()
  slide.background = { color: theme.background.replace('#', '') }

  const metrics = getLayoutGuides(layout)
  slide.addText(`${employee.name} · 其他检查附件`, {
    x: metrics.sideMargin,
    y: metrics.topMargin,
    w: metrics.contentWidth,
    fontSize: layout.orientation === 'portrait' ? 26 : 22,
    bold: true,
    color: theme.primary.replace('#', ''),
  })

  const items = documents.map((item) => ({
    text: `${formatAttachmentType(item.type)}｜${item.label}｜${item.fileName}`,
    options: {
      bullet: true,
      fontSize: 18,
      color: theme.textDark.replace('#', ''),
    },
  }))

  slide.addText(items, {
    x: metrics.sideMargin,
    y: metrics.topMargin + 0.8,
    w: metrics.contentWidth,
    lineSpacingMultiple: 1.3,
  })

  slide.addText('注：PDF/其他资料请在 data 目录中查看原文件。', {
    x: metrics.sideMargin,
    y: layout.height - metrics.bottomMargin,
    fontSize: 14,
    color: theme.neutral.replace('#', ''),
  })
}

function formatAttachmentType (type) {
  if (type === 'image') return '影像'
  if (type === 'pdf') return 'PDF'
  return '其他'
}

function formatSummaryTeaser (summary) {
  if (!summary) {
    return '总结：未找到对应文档，请补充。'
  }

  const normalized = summary.replace(/\n+/g, ' ').trim()
  if (normalized.length <= 120) {
    return `总结概述：${normalized}`
  }
  return `总结概述：${normalized.slice(0, 117)}...`
}

function chunkText (text, chunkSize) {
  if (!text) return []
  const paragraphs = text.split('\n').filter(Boolean)
  const chunks = []

  paragraphs.forEach((paragraph) => {
    if (paragraph.length <= chunkSize) {
      chunks.push(paragraph)
    } else {
      for (let i = 0; i < paragraph.length; i += chunkSize) {
        chunks.push(paragraph.slice(i, i + chunkSize))
      }
    }
  })

  return chunks.length ? chunks : ['暂无内容']
}

function buildReportFileName (employee) {
  const safeName = sanitizeForFilename(`${employee.name}_${employee.id}`)
  return `员工体检报告_${safeName}_${formatDate(new Date())}.pptx`
}

module.exports = {
  initializePresentation,
  addOverviewSlide,
  addSummarySlide,
  addImageSlides,
  addDocumentSlide,
  buildReportFileName,
};

