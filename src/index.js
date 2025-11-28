const path = require('path')
const fs = require('fs-extra')
const {
  DATA_DIR,
  OUTPUT_DIR,
  PDF_IMAGE_DIR,
  PDF_TEMP_DIR,
  TEMPLATE_CANDIDATES,
  EMPLOYEE_SHEET_CANDIDATES,
} = require('./config')
const { resolveExistingPath } = require('./utils')
const { loadTemplateLayout, loadThemeColors, loadEmployees } = require('./loaders')
const { safeReadDir, collectEmployeeAssets, buildImageItems } = require('./assets')
const {
  initializePresentation,
  addSummarySlide,
  addImageSlides,
  buildReportFileName,
} = require('./presentation')
const { insertTemplateSlides } = require('./templateSlides')

/**
 * 主入口：读取员工资料并批量生成 PPT 报告。
 */
async function main () {
  // 1) 准备输出目录，避免中途写文件失败。
  await fs.ensureDir(OUTPUT_DIR)
  await fs.ensureDir(PDF_IMAGE_DIR)
  await fs.ensureDir(PDF_TEMP_DIR)

  // 2) 定位模板与员工表，支持备用文件名。
  const templatePath = await resolveExistingPath(TEMPLATE_CANDIDATES, '模板文件')
  const sheetPath = await resolveExistingPath(EMPLOYEE_SHEET_CANDIDATES, '员工表')

  // 3) 读取模板布局、配色以及员工数据。
  const layout = await loadTemplateLayout(templatePath)
  const theme = await loadThemeColors(templatePath)
  const employees = await loadEmployees(sheetPath)

  if (!employees.length) {
    console.warn('⚠️ 员工表为空，已结束。')
    return
  }

  const availableFiles = await safeReadDir(DATA_DIR)
  const successReports = []
  const skippedEmployees = []

  // 4) 针对每位员工构建个性化报告。
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
        console.warn(`⚠️ ${employee.name} 未生成：缺少体检结果与AI总结`)
        continue
      }

      const pptx = initializePresentation(layout)

      const imageItems = await buildImageItems(assetInfo, employee)
      addImageSlides(pptx, employee, imageItems, theme, layout)

      // // 将体检总结放到最后
      // addSummarySlide(pptx, employee, assetInfo, theme, layout)

      const outputName = buildReportFileName(employee)
      const outputPath = path.join(OUTPUT_DIR, outputName)
      await pptx.writeFile({ fileName: outputPath })
      await insertTemplateSlides(templatePath, outputPath, { employee, date: new Date() })

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

  // 5) 输出汇总信息，便于快速排查。
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
});

