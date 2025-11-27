const fs = require('fs-extra');
const path = require('path');
const { ROOT } = require('./config');

function normalizeName(value) {
  if (!value) return '';
  return String(value)
    .replace(/\.[^.]+$/, '')
    .replace(/[-_].*$/, '')
    .replace(/\s+/g, '')
    .toLowerCase();
}

function normalizeText(value) {
  if (value === undefined || value === null) return '';
  return String(value).trim();
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

function formatDate(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const hh = String(date.getHours()).padStart(2, '0');
  const mi = String(date.getMinutes()).padStart(2, '0');
  return `${yyyy}${mm}${dd}_${hh}${mi}`;
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

module.exports = {
  normalizeName,
  normalizeText,
  sanitizeForFilename,
  buildAsciiSafeLabel,
  formatDate,
  resolveExistingPath,
};

