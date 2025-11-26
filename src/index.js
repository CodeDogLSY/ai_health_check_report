const path = require('path');
const fs = require('fs-extra');
const ExcelJS = require('exceljs');
const mammoth = require('mammoth');
const PptxGenJS = require('pptxgenjs');
const PizZip = require('pizzip');
const { fromPath: pdfFromPath } = require('pdf2pic');
const crypto = require('crypto');
const { createCanvas } = require('canvas');
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');

pdfjsLib.GlobalWorkerOptions.workerSrc = require.resolve('pdfjs-dist/legacy/build/pdf.worker.js');

const ROOT = path.resolve(__dirname, '..');
const DATA_DIR = path.join(ROOT, 'data');
const OUTPUT_DIR = path.join(ROOT, 'output');
const PDF_IMAGE_DIR = path.join(OUTPUT_DIR, '_pdf_images');
const PDF_TEMP_DIR = path.join(PDF_IMAGE_DIR, '_tmp');
const TEMPLATE_CANDIDATES = ['2025员工体检报告（模板）.pptx', 'template.pptx'];
const EMPLOYEE_SHEET_CANDIDATES = ['员工表.xlsx', 'employees.xlsx'];

const IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.gif', '.bmp']);
const PDF_EXTENSIONS = new Set(['.pdf']);
const EMU_PER_INCH = 914400;

const DEFAULT_THEME = {
  primary: '#4472C4',
  secondary: '#ED7D31',
  neutral: '#A5A5A5',
  highlight: '#FFC000',
  textDark: '#000000',
  textLight: '#FFFFFF',
  background: '#F7F9FC',
};

const DEFAULT_LAYOUT = {
  width: 10,
  height: 5.625,
  orientation: 'landscape',
};

async function main() {
  await fs.ensureDir(OUTPUT_DIR);
  await fs.ensureDir(PDF_IMAGE_DIR);
  await fs.ensureDir(PDF_TEMP_DIR);

  const templatePath = await resolveExistingPath(TEMPLATE_CANDIDATES, '模板文件');
  const sheetPath = await resolveExistingPath(EMPLOYEE_SHEET_CANDIDATES, '员工表');

  const layout = await loadTemplateLayout(templatePath);
  const theme = await loadThemeColors(templatePath);
  const employees = await loadEmployees(sheetPath);

  if (!employees.length) {
    console.warn('⚠️ 员工表为空，已结束。');
    return;
  }

  const availableFiles = await safeReadDir(DATA_DIR);
  const successReports = [];
  const skippedEmployees = [];

  for (const employee of employees) {
    try {
      const assetInfo = await collectEmployeeAssets(employee, availableFiles);
      const hasSummary = Boolean(assetInfo.summaryText.trim());
      const hasAttachments = assetInfo.attachments.length > 0;

      if (!hasSummary && !hasAttachments) {
        skippedEmployees.push({
          employee,
          reason: '缺少体检结果与AI总结',
        });
        console.warn(`⚠️ ${employee.name} 未生成：缺少体检结果与AI总结`);
        continue;
      }

      const pptx = initializePresentation(layout);
      addCoverSlide(pptx, employee, assetInfo, theme, layout);
      addOverviewSlide(pptx, employee, assetInfo, theme, layout);
      addSummarySlide(pptx, employee, assetInfo, theme, layout);

      const imageItems = await buildImageItems(assetInfo, employee);
      addImageSlides(pptx, employee, imageItems, theme, layout);
      addDocumentSlide(pptx, employee, assetInfo.attachments, theme, layout);

      const outputName = buildReportFileName(employee);
      const outputPath = path.join(OUTPUT_DIR, outputName);
      await pptx.writeFile({ fileName: outputPath });

      successReports.push({ employee, outputPath });
      console.log(`✓ 已生成 ${employee.name}（${employee.id}）：${outputPath}`);
    } catch (error) {
      skippedEmployees.push({
        employee,
        reason: `生成失败：${error.message}`,
      });
      console.error(`❌ ${employee.name} 生成失败：${error.message}`);
    }
  }

  console.log('\n===== 生成统计 =====');
  console.log(`✅ 已生成：${successReports.length} 人`);
  successReports.forEach((item) => {
    console.log(`  - ${item.employee.name}（${item.employee.id}） -> ${item.outputPath}`);
  });

  console.log(`⚠️ 未生成：${skippedEmployees.length} 人`);
  skippedEmployees.forEach((item) => {
    console.log(`  - ${item.employee.name}（${item.employee.id}）：${item.reason}`);
  });
}

async function loadThemeColors(templatePath) {
  if (!(await fs.pathExists(templatePath))) {
    console.warn('⚠️ 未找到模板文件，使用默认颜色。');
    return DEFAULT_THEME;
  }

  try {
    const buffer = await fs.readFile(templatePath);
    const zip = new PizZip(buffer);
    const themeFile = zip.file('ppt/theme/theme1.xml');

    if (!themeFile) {
      return DEFAULT_THEME;
    }

    const xml = themeFile.asText();
    const colors = {
      accent1: extractHex(xml, 'accent1'),
      accent2: extractHex(xml, 'accent2'),
      accent3: extractHex(xml, 'accent3'),
      accent4: extractHex(xml, 'accent4'),
      accent5: extractHex(xml, 'accent5'),
      accent6: extractHex(xml, 'accent6'),
      dk1: extractHex(xml, 'dk1'),
      lt1: extractHex(xml, 'lt1'),
      lt2: extractHex(xml, 'lt2'),
    };

    return {
      primary: buildColor(colors.accent1, DEFAULT_THEME.primary),
      secondary: buildColor(colors.accent2, DEFAULT_THEME.secondary),
      neutral: buildColor(colors.accent3, DEFAULT_THEME.neutral),
      highlight: buildColor(colors.accent4, DEFAULT_THEME.highlight),
      accent5: buildColor(colors.accent5, '#5B9BD5'),
      accent6: buildColor(colors.accent6, '#70AD47'),
      textDark: buildColor(colors.dk1, DEFAULT_THEME.textDark),
      textLight: buildColor(colors.lt1, DEFAULT_THEME.textLight),
      background: buildColor(colors.lt2, DEFAULT_THEME.background),
    };
  } catch (error) {
    console.warn(`⚠️ 读取模板主题失败，将使用默认颜色。原因：${error.message}`);
    return DEFAULT_THEME;
  }
}

async function loadTemplateLayout(templatePath) {
  const fallback = { ...DEFAULT_LAYOUT };

  if (!(await fs.pathExists(templatePath))) {
    return fallback;
  }

  try {
    const buffer = await fs.readFile(templatePath);
    const zip = new PizZip(buffer);
    const presentationFile = zip.file('ppt/presentation.xml');

    if (!presentationFile) {
      return fallback;
    }

    const xml = presentationFile.asText();
    const sizeMatch = xml.match(/<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/i);

    if (!sizeMatch) {
      return fallback;
    }

    const widthEmu = parseInt(sizeMatch[1], 10);
    const heightEmu = parseInt(sizeMatch[2], 10);

    if (Number.isNaN(widthEmu) || Number.isNaN(heightEmu)) {
      return fallback;
    }

    const width = Number((widthEmu / EMU_PER_INCH).toFixed(3));
    const height = Number((heightEmu / EMU_PER_INCH).toFixed(3));
    const orientation = width >= height ? 'landscape' : 'portrait';

    return { width, height, orientation };
  } catch (error) {
    console.warn(`⚠️ 模板布局解析失败，将使用默认 16:9 布局。原因：${error.message}`);
    return fallback;
  }
}

function extractHex(xml, tag) {
  const srgbRegex = new RegExp(`<a:${tag}>[\\s\\S]*?<a:srgbClr[^>]*?val="([0-9A-F]{6})"`, 'i');
  const sysRegex = new RegExp(`<a:${tag}>[\\s\\S]*?<a:sysClr[^>]*?lastClr="([0-9A-F]{6})"`, 'i');

  const srgbMatch = xml.match(srgbRegex);
  if (srgbMatch) {
    return srgbMatch[1];
  }

  const sysMatch = xml.match(sysRegex);
  if (sysMatch) {
    return sysMatch[1];
  }

  return null;
}

function buildColor(value, fallback) {
  if (!value) {
    return fallback;
  }
  return `#${value.toUpperCase()}`;
}

async function loadEmployees(sheetPath) {
  if (!(await fs.pathExists(sheetPath))) {
    throw new Error(`未找到员工表：${sheetPath}`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(sheetPath);
  const worksheet = workbook.worksheets[0];

  if (!worksheet) {
    return [];
  }

  const headerRow = worksheet.getRow(1);
  const headers = headerRow.values.map((value) => (typeof value === 'string' ? value.trim() : value));

  const employees = [];
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const record = {};
    row.values.forEach((cell, cellIndex) => {
      const header = headers[cellIndex];
      if (!header || typeof header !== 'string') return;
      record[header] = typeof cell === 'string' ? cell.trim() : cell;
    });

    const name = normalizeText(record['姓名'] || record['name']);
    if (!name) return;

    employees.push({
      name,
      id: normalizeText(record['工号'] || record['员工工号'] || record['编号']) || '未知',
      raw: record,
    });
  });

  return employees;
}

async function safeReadDir(dir) {
  if (!(await fs.pathExists(dir))) {
    console.warn(`⚠️ 数据目录不存在：${dir}`);
    return [];
  }
  return fs.readdir(dir);
}

async function collectEmployeeAssets(employee, files) {
  const normalized = normalizeName(employee.name);
  const matchingFiles = files.filter((file) => normalizeName(file) === normalized);

  const summaryFile = matchingFiles.find((file) => file.includes('总结') && path.extname(file).toLowerCase() === '.docx');
  const summaryText = summaryFile
    ? await extractSummaryText(path.join(DATA_DIR, summaryFile))
    : '';

  const attachments = matchingFiles
    .filter((file) => file !== summaryFile)
    .map((file) => {
      const ext = path.extname(file).toLowerCase();
      const type = IMAGE_EXTENSIONS.has(ext) ? 'image' : PDF_EXTENSIONS.has(ext) ? 'pdf' : 'other';
      const label = buildAttachmentLabel(file);
      return {
        fileName: file,
        fullPath: path.join(DATA_DIR, file),
        label,
        type,
        ext,
      };
    });

  return {
    employee,
    summaryFile: summaryFile ? path.join(DATA_DIR, summaryFile) : null,
    summaryText: summaryText || '',
    attachments,
    stats: {
      totalFiles: matchingFiles.length,
      images: attachments.filter((item) => item.type === 'image').length,
      pdfs: attachments.filter((item) => item.type === 'pdf').length,
      others: attachments.filter((item) => item.type === 'other').length,
    },
  };
}

async function extractSummaryText(filePath) {
  try {
    const { value } = await mammoth.extractRawText({ path: filePath });
    return value
      .split('\n')
      .map((item) => item.trim())
      .filter(Boolean)
      .join('\n');
  } catch (error) {
    console.warn(`⚠️ 无法读取AI总结：${filePath}，原因：${error.message}`);
    return '';
  }
}

function buildAttachmentLabel(fileName) {
  const base = path.basename(fileName, path.extname(fileName));
  const parts = base.split(/[-_]/);
  return parts.slice(1).join('-') || '附件';
}

function initializePresentation(layout) {
  const pptx = new PptxGenJS();
  const layoutName = `TEMPLATE_${layout.width}x${layout.height}`;
  pptx.defineLayout({
    name: layoutName,
    width: layout.width,
    height: layout.height,
  });
  pptx.layout = layoutName;
  pptx.author = 'Health Manage AI';
  pptx.company = 'Health Manage AI';
  pptx.subject = '员工体检报告';
  pptx.title = '2025 员工体检报告';
  return pptx;
}

function getLayoutGuides(layout) {
  const sideMargin = Number((layout.width * 0.08).toFixed(3));
  const topMargin = Number((layout.height * 0.05).toFixed(3));
  const bottomMargin = Number((layout.height * 0.06).toFixed(3));
  const contentWidth = Number((layout.width - sideMargin * 2).toFixed(3));
  return { sideMargin, topMargin, bottomMargin, contentWidth };
}

function addCoverSlide(pptx, employee, assets, theme, layout) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.background.replace('#', '') };

  const metrics = getLayoutGuides(layout);
  const headerHeight = Math.max(0.4, layout.height * 0.04);

  slide.addShape('rect', {
    x: 0,
    y: 0,
    w: layout.width,
    h: headerHeight,
    fill: { color: theme.primary.replace('#', '') },
    line: { color: theme.primary.replace('#', '') },
  });

  slide.addText('2025 员工体检报告', {
    x: metrics.sideMargin,
    y: headerHeight * 0.2,
    w: metrics.contentWidth,
    fontSize: layout.orientation === 'portrait' ? 24 : 20,
    color: theme.textLight.replace('#', ''),
  });

  const heroStartY = headerHeight + metrics.topMargin;
  slide.addText(employee.name, {
    x: metrics.sideMargin,
    y: heroStartY,
    w: metrics.contentWidth,
    fontSize: layout.orientation === 'portrait' ? 48 : 40,
    color: theme.primary.replace('#', ''),
    bold: true,
  });

  slide.addText(`工号：${employee.id}`, {
    x: metrics.sideMargin,
    y: heroStartY + 0.9,
    w: metrics.contentWidth,
    fontSize: layout.orientation === 'portrait' ? 22 : 18,
    color: theme.textDark.replace('#', ''),
  });

  slide.addText(`数据文件：${assets.stats.totalFiles}`, {
    x: metrics.sideMargin,
    y: heroStartY + 1.5,
    fontSize: layout.orientation === 'portrait' ? 20 : 18,
    color: theme.secondary.replace('#', ''),
  });

  const summaryY = heroStartY + (layout.orientation === 'portrait' ? 2.2 : 1.8);
  const summaryHeight = Math.max(2, layout.height - summaryY - metrics.bottomMargin);
  slide.addText(formatSummaryTeaser(assets.summaryText), {
    x: metrics.sideMargin,
    y: summaryY,
    w: metrics.contentWidth,
    h: summaryHeight,
    fontSize: 16,
    color: theme.textDark.replace('#', ''),
    lineSpacingMultiple: 1.2,
  });
}

function addOverviewSlide(pptx, employee, assets, theme, layout) {
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };

  const metrics = getLayoutGuides(layout);
  slide.addText(`${employee.name} · 检查项概览`, {
    x: metrics.sideMargin,
    y: metrics.topMargin,
    w: metrics.contentWidth,
    fontSize: layout.orientation === 'portrait' ? 28 : 24,
    bold: true,
    color: theme.primary.replace('#', ''),
  });

  const headerCells = [
    { text: '类别', options: { bold: true, color: theme.textLight.replace('#', ''), fill: theme.primary.replace('#', '') } },
    { text: '文件名', options: { bold: true, color: theme.textLight.replace('#', ''), fill: theme.primary.replace('#', '') } },
    { text: '备注', options: { bold: true, color: theme.textLight.replace('#', ''), fill: theme.primary.replace('#', '') } },
  ];

  const rows = [headerCells];

  if (!assets.attachments.length) {
    rows.push([
      { text: '暂无', options: {} },
      { text: '未检测到检查文件', options: {} },
      { text: '', options: {} },
    ]);
  } else {
    assets.attachments.forEach((item) => {
      rows.push([
        { text: formatAttachmentType(item.type), options: {} },
        { text: item.fileName, options: {} },
        { text: item.label || '附件', options: {} },
      ]);
    });
  }

  const tableY = metrics.topMargin + 0.8;
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
  });
}

function addSummarySlide(pptx, employee, assets, theme, layout) {
  const slide = pptx.addSlide();
  slide.background = { color: theme.background.replace('#', '') };

  const metrics = getLayoutGuides(layout);
  slide.addText(`${employee.name} · AI 体检总结`, {
    x: metrics.sideMargin,
    y: metrics.topMargin,
    w: metrics.contentWidth,
    fontSize: layout.orientation === 'portrait' ? 28 : 24,
    bold: true,
    color: theme.primary.replace('#', ''),
  });

  const summary = assets.summaryText || '暂无AI总结，请补充。';
  const paragraphs = chunkText(summary, 120);
  const textRuns = paragraphs.map((paragraph) => ({
    text: paragraph,
    options: {
      bullet: true,
      color: theme.textDark.replace('#', ''),
      fontSize: 18,
    },
  }));

  slide.addText(textRuns, {
    x: metrics.sideMargin,
    y: metrics.topMargin + 0.8,
    w: metrics.contentWidth,
    h: layout.height - metrics.topMargin - metrics.bottomMargin - 1,
    lineSpacingMultiple: 1.2,
  });
}

function addImageSlides(pptx, employee, imageItems, theme, layout) {
  if (!imageItems.length) {
    return;
  }

  const metrics = getLayoutGuides(layout);
  imageItems.forEach((image, index) => {
    const slide = pptx.addSlide();
    slide.background = { color: theme.background.replace('#', '') };

    slide.addText(`${employee.name} · 影像资料 ${index + 1}/${imageItems.length}`, {
      x: metrics.sideMargin,
      y: metrics.topMargin,
      w: metrics.contentWidth,
      fontSize: layout.orientation === 'portrait' ? 26 : 22,
      bold: true,
      color: theme.primary.replace('#', ''),
    });

    slide.addText(image.label, {
      x: metrics.sideMargin,
      y: metrics.topMargin + 0.6,
      w: metrics.contentWidth,
      fontSize: layout.orientation === 'portrait' ? 20 : 18,
      color: theme.secondary.replace('#', ''),
    });

    const imagePayload = image.data
      ? { data: image.data }
      : { path: image.fullPath };

    const imageTop = metrics.topMargin + 1.1;
    const maxWidth = metrics.contentWidth;
    const maxHeight = Math.max(3, layout.height - imageTop - metrics.bottomMargin);

    slide.addImage({
      ...imagePayload,
      x: metrics.sideMargin,
      y: imageTop,
      w: maxWidth,
      h: maxHeight,
      sizing: { type: 'contain', w: maxWidth, h: maxHeight },
    });
  });
}

function addDocumentSlide(pptx, employee, attachments, theme, layout) {
  const documents = attachments.filter((item) => item.type !== 'image');
  if (!documents.length) {
    return;
  }

  const slide = pptx.addSlide();
  slide.background = { color: theme.background.replace('#', '') };

  const metrics = getLayoutGuides(layout);
  slide.addText(`${employee.name} · 其他检查附件`, {
    x: metrics.sideMargin,
    y: metrics.topMargin,
    w: metrics.contentWidth,
    fontSize: layout.orientation === 'portrait' ? 26 : 22,
    bold: true,
    color: theme.primary.replace('#', ''),
  });

  const items = documents.map((item) => ({
    text: `${formatAttachmentType(item.type)}｜${item.label}｜${item.fileName}`,
    options: {
      bullet: true,
      fontSize: 18,
      color: theme.textDark.replace('#', ''),
    },
  }));

  slide.addText(items, {
    x: metrics.sideMargin,
    y: metrics.topMargin + 0.8,
    w: metrics.contentWidth,
    lineSpacingMultiple: 1.3,
  });

  slide.addText('注：PDF/其他资料请在 data 目录中查看原文件。', {
    x: metrics.sideMargin,
    y: layout.height - metrics.bottomMargin,
    fontSize: 14,
    color: theme.neutral.replace('#', ''),
  });
}

function formatAttachmentType(type) {
  if (type === 'image') return '影像';
  if (type === 'pdf') return 'PDF';
  return '其他';
}

function formatSummaryTeaser(summary) {
  if (!summary) {
    return 'AI总结：未找到对应文档，请补充。';
  }

  const normalized = summary.replace(/\n+/g, ' ').trim();
  if (normalized.length <= 120) {
    return `AI总结概述：${normalized}`;
  }
  return `AI总结概述：${normalized.slice(0, 117)}...`;
}

function chunkText(text, chunkSize) {
  if (!text) return [];
  const paragraphs = text.split('\n').filter(Boolean);
  const chunks = [];

  paragraphs.forEach((paragraph) => {
    if (paragraph.length <= chunkSize) {
      chunks.push(paragraph);
    } else {
      for (let i = 0; i < paragraph.length; i += chunkSize) {
        chunks.push(paragraph.slice(i, i + chunkSize));
      }
    }
  });

  return chunks.length ? chunks : ['暂无内容'];
}

function normalizeName(value) {
  if (!value) return '';
  return String(value)
    .replace(/\.[^.]+$/, '')
    .replace(/[-_].*$/, '')
    .replace(/\s+/g, '')
    .toLowerCase();
}

async function resolveExistingPath(candidates, label) {
  for (const candidate of candidates) {
    const fullPath = path.join(ROOT, candidate);
    if (await fs.pathExists(fullPath)) {
      if (candidate !== candidates[0]) {
        console.warn(`ℹ️ ${label}使用备用路径：${candidate}`);
      }
      return fullPath;
    }
  }
  throw new Error(`${label}缺失，请确认以下文件之一存在：${candidates.join(', ')}`);
}

function normalizeText(value) {
  if (value === undefined || value === null) return '';
  return String(value).trim();
}

function formatDate(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const hh = String(date.getHours()).padStart(2, '0');
  const mi = String(date.getMinutes()).padStart(2, '0');
  return `${yyyy}${mm}${dd}_${hh}${mi}`;
}

async function buildImageItems(assetInfo, employee) {
  const imageItems = assetInfo.attachments
    .filter((item) => item.type === 'image')
    .map((item) => ({
      label: item.label,
      fullPath: item.fullPath,
    }));

  const pdfAttachments = assetInfo.attachments.filter((item) => item.type === 'pdf');
  for (const pdf of pdfAttachments) {
    const converted = await convertPdfAttachment(pdf, employee);
    imageItems.push(...converted);
  }

  return imageItems;
}

async function convertPdfAttachment(pdfAttachment, employee) {
  let tempPdfPath = '';
  try {
    const safeLabel = buildAsciiSafeLabel(`${employee.id}_${employee.name}_${pdfAttachment.label}`);
    const tempPdfName = `${crypto.randomUUID()}.pdf`;
    tempPdfPath = path.join(PDF_TEMP_DIR, tempPdfName);
    await fs.copyFile(pdfAttachment.fullPath, tempPdfPath);

    const pdfImgPages = await convertWithPdfRenderer(tempPdfPath, pdfAttachment);
    if (pdfImgPages.length) {
      return pdfImgPages;
    }

    const pdf2picPages = await convertWithPdf2Pic(tempPdfPath, pdfAttachment, safeLabel);
    return pdf2picPages;
  } catch (error) {
    console.warn(`⚠️ PDF 转图失败（${pdfAttachment.fileName}）：${error.message}`);
    return [];
  } finally {
    if (tempPdfPath) {
      await fs.remove(tempPdfPath).catch(() => {});
    }
  }
}

function buildReportFileName(employee) {
  const safeName = sanitizeForFilename(`${employee.name}_${employee.id}`);
  return `员工体检报告_${safeName}_${formatDate(new Date())}.pptx`;
}

function sanitizeForFilename(value) {
  return value.replace(/[<>:"/\\|?*]/g, '').replace(/\s+/g, '');
}

function buildAsciiSafeLabel(value) {
  const sanitized = sanitizeForFilename(value).replace(/[^\x00-\x7F]/g, '');
  if (sanitized) {
    return sanitized;
  }
  return `pdf_${Date.now()}`;
}

async function convertWithPdfRenderer(pdfPath, attachment) {
  try {
    const pdfBuffer = await fs.readFile(pdfPath);
    const pdfData = new Uint8Array(pdfBuffer);
    const pdfDoc = await pdfjsLib.getDocument({ data: pdfData, verbosity: 0 }).promise;
    const pages = [];

    for (let pageNumber = 1; pageNumber <= pdfDoc.numPages; pageNumber++) {
      const page = await pdfDoc.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 2 });
      const canvas = createCanvas(viewport.width, viewport.height);
      const context = canvas.getContext('2d');

      await page.render({ canvasContext: context, viewport }).promise;
      const imageBuffer = canvas.toBuffer('image/png');
      pages.push({
        label: `${attachment.label} 第${pageNumber}页`,
        data: `data:image/png;base64,${imageBuffer.toString('base64')}`,
      });
    }

    return pages;
  } catch (error) {
    console.warn(`⚠️ 内置 PDF 渲染失败（${attachment.fileName}）：${error.message}`);
    return [];
  }
}

async function convertWithPdf2Pic(pdfPath, attachment, safeLabel) {
  try {
    const convert = pdfFromPath(pdfPath, {
      density: 144,
      format: 'png',
      width: 1200,
      height: 800,
      saveFilename: safeLabel,
      savePath: PDF_IMAGE_DIR,
    });

    const pages = await convert.bulk(-1, { responseType: 'base64' });
    return pages.map((pageResult, idx) => ({
      label: `${attachment.label} 第${pageResult.page || idx + 1}页`,
      data: `data:image/png;base64,${pageResult.base64}`,
    }));
  } catch (error) {
    console.warn(`⚠️ pdf2pic 转换失败（${attachment.fileName}）：${error.message}`);
    return [];
  }
}

main().catch((error) => {
  console.error('❌ 生成失败：', error);
  process.exitCode = 1;
});

