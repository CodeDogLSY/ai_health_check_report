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

    slide.addText(image.label, {
      x: metrics.sideMargin,
      y: metrics.topMargin,
      w: metrics.contentWidth,
      fontSize: layout.orientation === 'portrait' ? 26 : 22,
      bold: true,
      color: theme.primary.replace('#', ''),
    })

    const imagePayload = image.data
      ? { data: image.data }
      : { path: image.fullPath }

    const imageTop = metrics.topMargin + 0.4
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
  addSummarySlide,
  addImageSlides,
  buildReportFileName,
};

