const path = require('path');

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

module.exports = {
  ROOT,
  DATA_DIR,
  OUTPUT_DIR,
  PDF_IMAGE_DIR,
  PDF_TEMP_DIR,
  TEMPLATE_CANDIDATES,
  EMPLOYEE_SHEET_CANDIDATES,
  IMAGE_EXTENSIONS,
  PDF_EXTENSIONS,
  EMU_PER_INCH,
  DEFAULT_THEME,
  DEFAULT_LAYOUT,
};

