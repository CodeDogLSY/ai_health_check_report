const fs = require('fs-extra');
const ExcelJS = require('exceljs');
const PizZip = require('pizzip');
const { DEFAULT_THEME, DEFAULT_LAYOUT, EMU_PER_INCH } = require('./config');
const { normalizeText } = require('./utils');

/**
 * 从 PPT 模板中提取主题配色，找不到则返回默认主题。
 */
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

/**
 * 解析 PPT 模板的宽高布局信息。
 */
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

/**
 * 加载 Excel 员工表，并抽取姓名与工号。
 */
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

module.exports = {
  loadThemeColors,
  loadTemplateLayout,
  loadEmployees,
};

