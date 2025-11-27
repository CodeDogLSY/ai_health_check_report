const fs = require('fs-extra');
const path = require('path');
const PizZip = require('pizzip');

const POSIX = path.posix;

const CONTENT_TYPES = {
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  slideLayout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  slideMaster: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
};

const REL_TYPES = {
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
};

async function insertTemplateSlides(templatePath, outputPath, options = {}) {
  const [templateBuffer, outputBuffer] = await Promise.all([
    fs.readFile(templatePath),
    fs.readFile(outputPath),
  ]);

  const templateZip = new PizZip(templateBuffer);
  const outputZip = new PizZip(outputBuffer);

  const slidesToCopy = [
    { templateSlide: 1, position: 'start' },
    { templateSlide: 4, position: 'end' },
  ];

  const state = initializeState(outputZip, options);

  for (const config of slidesToCopy) {
    copyTemplateSlide({
      templateZip,
      outputZip,
      templateSlideNumber: config.templateSlide,
      position: config.position,
      state,
      options,
    });
  }

  if (!state.newSlides.length) {
    return;
  }

  updatePresentationDocuments(outputZip, state);

  const updatedBuffer = outputZip.generate({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  await fs.writeFile(outputPath, updatedBuffer);
}

function initializeState(zip, options) {
  const presentationXml = zip.file('ppt/presentation.xml');
  const presentationRels = zip.file('ppt/_rels/presentation.xml.rels');
  const contentTypes = zip.file('[Content_Types].xml');

  if (!presentationXml || !presentationRels || !contentTypes) {
    throw new Error('目标 PPT 结构不完整，缺少 presentation 元数据');
  }

  const presContent = presentationXml.asText();
  const presRelsContent = presentationRels.asText();

  const slideIdMatches = [...presContent.matchAll(/<p:sldId[^>]*id="(\d+)"/g)];
  const masterIdMatches = [...presContent.matchAll(/<p:sldMasterId[^>]*id="(\d+)"/g)];
  const relIdMatches = [...presRelsContent.matchAll(/Id="rId(\d+)"/g)];

  const slideFiles = zip.file(/^ppt\/slides\/slide\d+\.xml$/i);
  const layoutFiles = zip.file(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/i);
  const masterFiles = zip.file(/^ppt\/slideMasters\/slideMaster\d+\.xml$/i);
  const themeFiles = zip.file(/^ppt\/theme\/theme\d+\.xml$/i);

  return {
    presentationXml,
    presContent,
    presentationRels,
    presRelsContent,
    contentTypes,
    contentTypesContent: contentTypes.asText(),
    nextSlideNumber: getNextNumber(slideFiles, /slide(\d+)\.xml/i),
    nextSlideId: slideIdMatches.length
      ? Math.max(...slideIdMatches.map((m) => parseInt(m[1], 10)))
      : 256,
    nextRelId: relIdMatches.length ? Math.max(...relIdMatches.map((m) => parseInt(m[1], 10))) : 10,
    nextLayoutNumber: getNextNumber(layoutFiles, /slideLayout(\d+)\.xml/i),
    nextMasterNumber: getNextNumber(masterFiles, /slideMaster(\d+)\.xml/i),
    nextThemeNumber: getNextNumber(themeFiles, /theme(\d+)\.xml/i),
    nextMasterId: masterIdMatches.length
      ? Math.max(...masterIdMatches.map((m) => parseInt(m[1], 10)))
      : 100,
    mediaCounter: getInitialMediaCounter(zip),
    layoutMap: new Map(),
    masterMap: new Map(),
    themeMap: new Map(),
    mediaMap: new Map(),
    newSlides: [],
    newMasters: new Map(),
    newMastersList: [],
    options,
  };
}

function getNextNumber(files, regex) {
  if (!files || !files.length) {
    return 1;
  }
  const numbers = files
    .map((file) => {
      const match = file.name.match(regex);
      return match ? parseInt(match[1], 10) : 0;
    })
    .filter((num) => Number.isFinite(num));
  if (!numbers.length) {
    return 1;
  }
  return Math.max(...numbers) + 1;
}

function copyTemplateSlide({ templateZip, outputZip, templateSlideNumber, position, state, options }) {
  const templateSlidePath = `ppt/slides/slide${templateSlideNumber}.xml`;
  const templateSlide = templateZip.file(templateSlidePath);
  if (!templateSlide) {
    console.warn(`⚠️ 模板缺少第${templateSlideNumber}页，跳过注入`);
    return;
  }

  const newSlideNumber = state.nextSlideNumber++;
  const newSlidePath = `ppt/slides/slide${newSlideNumber}.xml`;
  let slideXml = templateSlide.asText();

  if (position === 'start') {
    slideXml = applyCoverPlaceholders(slideXml, options.employee, options.date);
  }

  outputZip.file(newSlidePath, slideXml);

  cloneSlideRelationships({
    templateZip,
    outputZip,
    templateSlideNumber,
    newSlideNumber,
    state,
  });

  ensureContentType(state, `/ppt/slides/slide${newSlideNumber}.xml`, CONTENT_TYPES.slide);

  state.newSlides.push({
    slideNumber: newSlideNumber,
    position,
  });
}

function cloneSlideRelationships({ templateZip, outputZip, templateSlideNumber, newSlideNumber, state }) {
  const relsPath = `ppt/slides/_rels/slide${templateSlideNumber}.xml.rels`;
  const templateRels = templateZip.file(relsPath);
  const newRelsPath = `ppt/slides/_rels/slide${newSlideNumber}.xml.rels`;

  if (!templateRels) {
    outputZip.file(newRelsPath, buildEmptyRelationships());
    return;
  }

  const relItems = parseRelationships(templateRels.asText());
  const updated = relItems.map((rel) => {
    if (rel.Type === REL_TYPES.slideLayout) {
      const absolutePath = resolveRelationshipPath(`ppt/slides/slide${templateSlideNumber}.xml`, rel.Target);
      const layoutInfo = ensureLayoutAssets({
        templateZip,
        outputZip,
        templateLayoutPath: absolutePath,
        state,
      });

      rel.Target = relativePath(`ppt/slides/slide${newSlideNumber}.xml`, layoutInfo.newPath);
      return rel;
    }

    if (rel.Type === REL_TYPES.image) {
      const absoluteMediaPath = resolveRelationshipPath(`ppt/slides/slide${templateSlideNumber}.xml`, rel.Target);
      const newMediaTarget = copyMediaFile({
        templateZip,
        outputZip,
        sourcePath: absoluteMediaPath,
        state,
      });
      rel.Target = relativePath(`ppt/slides/slide${newSlideNumber}.xml`, newMediaTarget);
      return rel;
    }

    // 其他引用（如 notes）直接复制
    const absoluteTarget = resolveRelationshipPath(`ppt/slides/slide${templateSlideNumber}.xml`, rel.Target);
    const newTarget = copyGenericPart({
      templateZip,
      outputZip,
      sourcePath: absoluteTarget,
      state,
    });
    rel.Target = relativePath(`ppt/slides/slide${newSlideNumber}.xml`, newTarget);
    return rel;
  });

  outputZip.file(newRelsPath, buildRelationshipsXml(updated));
}

function ensureLayoutAssets({ templateZip, outputZip, templateLayoutPath, state }) {
  if (state.layoutMap.has(templateLayoutPath)) {
    return state.layoutMap.get(templateLayoutPath);
  }

  // 通过解析布局的关系找到对应母版，确保母版先被复制
  const relsPath = layoutRelationshipsPath(templateLayoutPath);
  const relFile = templateZip.file(relsPath);
  if (!relFile) {
    throw new Error(`模板布局缺少关系文件：${relsPath}`);
  }

  const relItems = parseRelationships(relFile.asText());
  const masterRel = relItems.find((rel) => rel.Type === REL_TYPES.slideMaster);
  if (!masterRel) {
    throw new Error(`布局文件未关联母版：${templateLayoutPath}`);
  }

  const masterPath = resolveRelationshipPath(templateLayoutPath, masterRel.Target);
  ensureMasterAssets({
    templateZip,
    outputZip,
    templateMasterPath: masterPath,
    state,
  });

  if (!state.layoutMap.has(templateLayoutPath)) {
    throw new Error(`复制布局失败：${templateLayoutPath}`);
  }

  return state.layoutMap.get(templateLayoutPath);
}

function ensureMasterAssets({ templateZip, outputZip, templateMasterPath, state }) {
  if (state.masterMap.has(templateMasterPath)) {
    return state.masterMap.get(templateMasterPath);
  }

  const masterFile = templateZip.file(templateMasterPath);
  if (!masterFile) {
    throw new Error(`模板缺少母版文件：${templateMasterPath}`);
  }

  const newMasterNumber = state.nextMasterNumber++;
  const newMasterPath = `ppt/slideMasters/slideMaster${newMasterNumber}.xml`;
  outputZip.file(newMasterPath, masterFile.asText());
  ensureContentType(state, `/${newMasterPath}`, CONTENT_TYPES.slideMaster);

  const masterRelsPath = masterRelationshipsPath(templateMasterPath);
  const templateRels = templateZip.file(masterRelsPath);
  const newMasterRelsPath = masterRelationshipsPath(newMasterPath);
  let themeInfo = null;

  if (templateRels) {
    const relItems = parseRelationships(templateRels.asText());
    const updated = relItems.map((rel) => {
      if (rel.Type === REL_TYPES.theme) {
        const themePath = resolveRelationshipPath(templateMasterPath, rel.Target);
        themeInfo = ensureThemeAssets({
          templateZip,
          outputZip,
          templateThemePath: themePath,
          state,
        });
        rel.Target = relativePath(newMasterPath, themeInfo.newPath);
        return rel;
      }

      if (rel.Type === REL_TYPES.slideLayout) {
        const layoutPath = resolveRelationshipPath(templateMasterPath, rel.Target);
        const layoutInfo = copyLayoutFromMaster({
          templateZip,
          outputZip,
          templateLayoutPath: layoutPath,
          newMasterPath,
          state,
        });
        rel.Target = relativePath(newMasterPath, layoutInfo.newPath);
        return rel;
      }

      const absolute = resolveRelationshipPath(templateMasterPath, rel.Target);
      const newTarget = copyGenericPart({
        templateZip,
        outputZip,
        sourcePath: absolute,
        state,
      });
      rel.Target = relativePath(newMasterPath, newTarget);
      return rel;
    });

    outputZip.file(newMasterRelsPath, buildRelationshipsXml(updated));
  } else {
    outputZip.file(newMasterRelsPath, buildEmptyRelationships());
  }

  const info = {
    newPath: newMasterPath,
    themePath: themeInfo ? themeInfo.newPath : null,
  };
  state.masterMap.set(templateMasterPath, info);
  state.newMasters.set(templateMasterPath, info);
  state.newMastersList.push(info);
  return info;
}

function copyLayoutFromMaster({ templateZip, outputZip, templateLayoutPath, newMasterPath, state }) {
  if (state.layoutMap.has(templateLayoutPath)) {
    return state.layoutMap.get(templateLayoutPath);
  }

  const layoutFile = templateZip.file(templateLayoutPath);
  if (!layoutFile) {
    throw new Error(`模板缺少布局文件：${templateLayoutPath}`);
  }

  const newLayoutNumber = state.nextLayoutNumber++;
  const newLayoutPath = `ppt/slideLayouts/slideLayout${newLayoutNumber}.xml`;
  outputZip.file(newLayoutPath, layoutFile.asText());
  ensureContentType(state, `/${newLayoutPath}`, CONTENT_TYPES.slideLayout);

  const relsPath = layoutRelationshipsPath(templateLayoutPath);
  const relFile = templateZip.file(relsPath);
  const newRelsPath = layoutRelationshipsPath(newLayoutPath);

  if (relFile) {
    const relItems = parseRelationships(relFile.asText());
    const updated = relItems.map((rel) => {
      if (rel.Type === REL_TYPES.slideMaster) {
        rel.Target = relativePath(newLayoutPath, newMasterPath);
        return rel;
      }

      if (rel.Type === REL_TYPES.image) {
        const mediaPath = resolveRelationshipPath(templateLayoutPath, rel.Target);
        const newMediaTarget = copyMediaFile({
          templateZip,
          outputZip,
          sourcePath: mediaPath,
          state,
        });
        rel.Target = relativePath(newLayoutPath, newMediaTarget);
        return rel;
      }

      const absolute = resolveRelationshipPath(templateLayoutPath, rel.Target);
      const newTarget = copyGenericPart({
        templateZip,
        outputZip,
        sourcePath: absolute,
        state,
      });
      rel.Target = relativePath(newLayoutPath, newTarget);
      return rel;
    });
    outputZip.file(newRelsPath, buildRelationshipsXml(updated));
  } else {
    outputZip.file(newRelsPath, buildEmptyRelationships());
  }

  const info = { newPath: newLayoutPath };
  state.layoutMap.set(templateLayoutPath, info);
  return info;
}

function ensureThemeAssets({ templateZip, outputZip, templateThemePath, state }) {
  if (state.themeMap.has(templateThemePath)) {
    return state.themeMap.get(templateThemePath);
  }

  const themeFile = templateZip.file(templateThemePath);
  if (!themeFile) {
    throw new Error(`模板缺少主题文件：${templateThemePath}`);
  }

  const newThemeNumber = state.nextThemeNumber++;
  const newThemePath = `ppt/theme/theme${newThemeNumber}.xml`;
  outputZip.file(newThemePath, themeFile.asText());
  ensureContentType(state, `/${newThemePath}`, CONTENT_TYPES.theme);

  const info = { newPath: newThemePath };
  state.themeMap.set(templateThemePath, info);
  return info;
}

function copyMediaFile({ templateZip, outputZip, sourcePath, state }) {
  if (state.mediaMap.has(sourcePath)) {
    return state.mediaMap.get(sourcePath);
  }

  const file = templateZip.file(sourcePath);
  if (!file) {
    throw new Error(`模板缺少资源：${sourcePath}`);
  }

  const ext = path.extname(sourcePath) || '.bin';
  const newName = `ppt/media/template_media_${state.mediaCounter++}${ext}`;
  outputZip.file(newName, file.asNodeBuffer());
  state.mediaMap.set(sourcePath, newName);
  return newName;
}

function copyGenericPart({ templateZip, outputZip, sourcePath, state }) {
  if (!sourcePath.startsWith('ppt/')) {
    return sourcePath;
  }
  if (outputZip.file(sourcePath)) {
    return sourcePath;
  }
  const part = templateZip.file(sourcePath);
  if (part) {
    if (part.options && part.options.binary) {
      outputZip.file(sourcePath, part.asNodeBuffer());
    } else {
      outputZip.file(sourcePath, part.asText());
    }
  }

  if (sourcePath.startsWith('ppt/notesSlides/')) {
    ensureContentType(
      state,
      `/${sourcePath}`,
      'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    );
  }
  return sourcePath;
}

function ensureContentType(state, partName, contentType) {
  if (state.contentTypesContent.includes(`PartName="${partName}"`)) {
    return;
  }
  const override = `  <Override PartName="${partName}" ContentType="${contentType}"/>`;
  state.contentTypesContent = state.contentTypesContent.replace('</Types>', `${override}\n</Types>`);
}

function updatePresentationDocuments(zip, state) {
  const slideEntries = extractExistingSlideEntries(state.presContent);
  const updatedSlideEntries = buildSlideEntries(slideEntries, state);
  state.presContent = state.presContent.replace(slideEntries.raw, updatedSlideEntries.xml);

  const updatedRels = appendSlideRelationships(state.presRelsContent, state);
  state.presRelsContent = updatedRels;

  const updatedMasters = appendMasterEntries(state.presContent, state);
  state.presContent = updatedMasters;

  zip.file('ppt/presentation.xml', state.presContent);
  zip.file('ppt/_rels/presentation.xml.rels', state.presRelsContent);
  zip.file('[Content_Types].xml', state.contentTypesContent);
}

function extractExistingSlideEntries(presContent) {
  const match = presContent.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
  if (!match) {
    throw new Error('presentation.xml 缺少 p:sldIdLst');
  }
  return { raw: match[0], inner: match[1] };
}

function buildSlideEntries(existing, state) {
  const existingEntries = existing.inner
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);

  const newSlides = state.newSlides.map((slide) => {
    state.nextSlideId += 1;
    state.nextRelId += 1;
    slide.slideId = state.nextSlideId;
    slide.relId = `rId${state.nextRelId}`;
    return slide;
  });

  const entryFor = (slide) => `    <p:sldId id="${slide.slideId}" r:id="${slide.relId}"/>`;

  const startEntries = newSlides.filter((s) => s.position === 'start').map(entryFor);
  const endEntries = newSlides.filter((s) => s.position === 'end').map(entryFor);

  const xml = `<p:sldIdLst>
${[...startEntries, ...existingEntries, ...endEntries].join('\n')}
</p:sldIdLst>`;

  return { xml };
}

function appendSlideRelationships(relsContent, state) {
  const insertIndex = relsContent.indexOf('</Relationships>');
  if (insertIndex === -1) {
    throw new Error('presentation.xml.rels 内容异常');
  }

  const slideRels = state.newSlides
    .map(
      (slide) =>
        `  <Relationship Id="${slide.relId}" Type="${REL_TYPES.slide}" Target="slides/slide${slide.slideNumber}.xml"/>`,
    )
    .join('\n');

  return `${relsContent.slice(0, insertIndex)}${slideRels}\n${relsContent.slice(insertIndex)}`;
}

function appendMasterEntries(presContent, state) {
  if (!state.newMastersList.length) {
    return presContent;
  }

  const relsContent = state.presRelsContent;
  const insertIndex = relsContent.indexOf('</Relationships>');

  if (insertIndex === -1) {
    throw new Error('presentation.xml.rels 缺少 Relationships 结束标签');
  }

  const masterListMatch = presContent.match(/<p:sldMasterIdLst>([\s\S]*?)<\/p:sldMasterIdLst>/);
  const existingEntries = masterListMatch
    ? masterListMatch[1]
        .split('\n')
        .map((line) => line.trim())
        .filter(Boolean)
    : [];

  const newMasterEntries = [];
  const newMasterRels = [];

  for (const info of state.newMastersList) {
    state.nextMasterId += 1;
    state.nextRelId += 1;
    const relId = `rId${state.nextRelId}`;
    const entry = `    <p:sldMasterId id="${state.nextMasterId}" r:id="${relId}"/>`;
    newMasterEntries.push(entry);
    newMasterRels.push(
      `  <Relationship Id="${relId}" Type="${REL_TYPES.slideMaster}" Target="${info.newPath.replace(
        'ppt/',
        '',
      )}"/>`,
    );
  }

  const newMasterList = `<p:sldMasterIdLst>
${[...existingEntries, ...newMasterEntries].join('\n')}
</p:sldMasterIdLst>`;

  let updatedPres = presContent;
  if (masterListMatch) {
    updatedPres = presContent.replace(masterListMatch[0], newMasterList);
  } else {
    updatedPres = presContent.replace(
      '</p:presentation>',
      `${newMasterList}\n</p:presentation>`,
    );
  }

  state.presRelsContent = `${relsContent.slice(
    0,
    insertIndex,
  )}${newMasterRels.join('\n')}\n${relsContent.slice(insertIndex)}`;

  return updatedPres;
}

function parseRelationships(xml) {
  const relRegex = /<Relationship\s+([^>]+?)\/>/g;
  const attrRegex = /([A-Za-z:]+)="([^"]*)"/g;
  const rels = [];
  let match;
  while ((match = relRegex.exec(xml))) {
    const attrs = {};
    let attrMatch;
    while ((attrMatch = attrRegex.exec(match[1]))) {
      attrs[attrMatch[1]] = attrMatch[2];
    }
    rels.push(attrs);
  }
  return rels;
}

function buildRelationshipsXml(rels) {
  const items = rels
    .map((rel) => {
      const attrs = Object.entries(rel)
        .map(([key, value]) => `${key}="${value}"`)
        .join(' ');
      return `  <Relationship ${attrs}/>`;
    })
    .join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${items}
</Relationships>`;
}

function buildEmptyRelationships() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;
}

function resolveRelationshipPath(from, target) {
  const baseDir = POSIX.dirname(from);
  return POSIX.normalize(POSIX.join(baseDir, target));
}

function relativePath(from, to) {
  const fromDir = POSIX.dirname(from);
  return POSIX.relative(fromDir, to);
}

function layoutRelationshipsPath(layoutPath) {
  const baseName = POSIX.basename(layoutPath);
  return `ppt/slideLayouts/_rels/${baseName}.rels`;
}

function masterRelationshipsPath(masterPath) {
  const baseName = POSIX.basename(masterPath);
  return `ppt/slideMasters/_rels/${baseName}.rels`;
}

function getInitialMediaCounter(zip) {
  const mediaFiles = zip.file(/^ppt\/media\/.+\.(png|jpg|jpeg|bmp|gif)$/i);
  if (!mediaFiles || !mediaFiles.length) {
    return 1;
  }
  const numbers = mediaFiles
    .map((file) => {
      const match = file.name.match(/(\d+)\.(png|jpg|jpeg|bmp|gif)$/i);
      return match ? parseInt(match[1], 10) : 0;
    })
    .filter((num) => Number.isFinite(num));
  return numbers.length ? Math.max(...numbers) + 1 : 1;
}

function applyCoverPlaceholders(xml, employee = {}, dateValue = new Date()) {
  const dateText = formatCnDate(dateValue);
  const replacements = {
    姓名: employee.name || '',
    性别: employee.gender || '',
    年龄: employee.age || '',
    日期: dateText,
    工号: employee.id || '',
  };

  return replaceTextInSlidePreferred(xml, replacements);
}

function formatCnDate(dateInput) {
  const date = dateInput instanceof Date ? dateInput : new Date(dateInput || Date.now());
  if (Number.isNaN(date.getTime())) {
    return '';
  }
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}年${month}月${day}日`;
}

function escapeXmlValue(value) {
  if (!value) return '';
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function replaceTextInSlidePreferred(xmlContent, replacements) {
  let result = xmlContent;
  const optionalSpaces = '(?:\\s|\\u00A0|&nbsp;|&#160;|&#xA0;)*';

  for (const [key, value] of Object.entries(replacements)) {
    if (value === undefined || value === null) continue;
    const valueStr = String(value);
    if (!valueStr && key !== '性别' && key !== '年龄') continue;

    const escapedValue = escapeXmlValue(valueStr);
    const keyPattern = escapeRegex(key);

    const placeholderPattern = new RegExp(`\\{\\{${optionalSpaces}${keyPattern}${optionalSpaces}\\}\\}${optionalSpaces}`, 'gi');
    result = result.replace(placeholderPattern, (match) => {
      const suffixMatch = match.match(/((?:\\s|\\u00A0|&nbsp;|&#160;|&#xA0;)*)$/i);
      const suffix = suffixMatch ? suffixMatch[1] : '';
      return `${escapedValue}${suffix}`;
    });

    const bracketPattern = new RegExp(`\\[${optionalSpaces}${keyPattern}${optionalSpaces}\\]`, 'gi');
    result = result.replace(bracketPattern, escapedValue);

    const textPattern = new RegExp(`(<a:t[^>]*>)${optionalSpaces}${keyPattern}${optionalSpaces}(</a:t>)`, 'gi');
    result = result.replace(textPattern, `$1${escapedValue}$2`);
  }

  return result;
}

function escapeRegex(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

module.exports = {
  insertTemplateSlides,
};

