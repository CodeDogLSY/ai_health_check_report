const ExcelJS = require('exceljs')
const fs = require('fs-extra')

async function checkEmployees () {
  try {
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile('员工表.xlsx')
    const worksheet = workbook.worksheets[0]

    // 获取表头
    const headerRow = worksheet.getRow(1)
    const headers = []
    
    // 遍历表头单元格，只处理有值的单元格
    for (let i = 1; i <= headerRow.cellCount; i++) {
      const cell = headerRow.getCell(i)
      const value = cell.value
      if (value) {
        headers.push({
          name: typeof value === 'string' ? value.trim() : value,
          index: i
        })
      }
    }
    
    // 找到姓名和证件号列
    const nameCol = headers.find(h => h.name === '姓名')
    const idCol = headers.find(h => h.name === '证件号')

    if (!nameCol || !idCol) {
      console.error('未找到姓名或证件号列')
      return
    }

    console.log('姓名,证件号')

    // 遍历所有行，查找姓名为王磊的员工
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        const name = row.getCell(nameCol.index).value
        const id = row.getCell(idCol.index).value
        if (name === '王磊') {
          console.log(`${name},${id}`)
        }
      }
    })
  } catch (error) {
    console.error('读取员工表时出错:', error.message)
  }
}

checkEmployees()