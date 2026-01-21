const path = require("path")
const fs = require("fs-extra")

/**
 * 读取send_data目录下的PDF文件，从文件名中提取姓名和证件号，
 * 调用外部接口获取企业微信ID，然后发送PDF文件到对应的企业微信账号
 */
async function sendWechatEeMsg () {
  const ROOT = path.resolve(__dirname, ".")
  const SEND_DATA_DIR = path.join(ROOT, "send_data")
  if (!(await fs.pathExists(SEND_DATA_DIR))) {
    console.error(`❌ send_data 文件夹不存在`)
    return
  }

  const files = await fs.readdir(SEND_DATA_DIR)
  const pdfFiles = files.filter(
    (file) => path.extname(file).toLowerCase() === ".pdf",
  )

  if (pdfFiles.length === 0) {
    console.log(`ℹ️ send_data 文件夹内没有PDF文件`)
    return
  }

  console.log(`📋 找到 ${pdfFiles.length} 个PDF文件，开始提取证件号...`)

  let successCount = 0
  let failCount = 0

  // 导入axios和form-data库
  const axios = require("axios")
  const FormData = require("form-data")

  // 定义发送POST请求获取企信ID的函数
  async function getQixinId (sfz) {
    try {
      // 使用axios发送POST请求，将sfz参数放在URL中
      const response = await axios.post(
        `http://wxsite.yinda.cn:5182/cajserver/pro/caj-renlizy/ZhiGong/nologin/getQixinIdBySfz?sfz=${encodeURIComponent(sfz)}`,
        {},
        {
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
          },
        },
      )
      return response.data
    } catch (error) {
      if (error.response) {
        throw new Error(
          `请求失败: ${error.response.status} ${error.response.statusText}`,
        )
      } else if (error.request) {
        throw new Error(`请求失败: 没有收到响应`)
      } else {
        throw new Error(`请求失败: ${error.message}`)
      }
    }
  }

  // 定义发送form-data请求上传文件的函数
  async function sendFileToUser (userId, fileFullPath, fileName) {
    try {
      // 创建form-data对象
      const formData = new FormData()
      // 添加文件
      formData.append("file", fs.createReadStream(fileFullPath), {
        filename: fileName,
        contentType: "application/pdf",
      })

      // 使用axios发送POST请求
      const response = await axios.post(
        `https://product.cajcare.com:5182/wechat/caj/sunflower/sendFileToUser?userId=${encodeURIComponent(userId)}`,
        formData,
        {
          headers: {
            ...formData.getHeaders(),
          },
        },
      )
      return response.data
    } catch (error) {
      if (error.response) {
        throw new Error(
          `发送文件失败: ${error.response.status} ${error.response.statusText}`,
        )
      } else if (error.request) {
        throw new Error(`发送文件失败: 没有收到响应`)
      } else {
        throw new Error(`发送文件失败: ${error.message}`)
      }
    }
  }

  for (const pdfFile of pdfFiles) {
    try {
      // 从文件名中提取证件号，命名规则：体检报告_姓名_证件号.pdf 或 体检报告_姓名_证件号_数字.pdf
      const idMatch = pdfFile.match(
        /^体检报告_([^_]+)_([\dXx]+)(?:_\d+)?\.pdf$/,
      )
      if (!idMatch) {
        console.warn(`⚠️ 文件名格式不符合要求：${pdfFile}`)
        failCount++
        continue
      }

      const name = idMatch[1]
      const id = idMatch[2]
      const fileFullPath = path.join(SEND_DATA_DIR, pdfFile)
      console.log(`✅ ${pdfFile} -> 姓名：${name}，证件号：${id}`)

      // 调用接口获取企信ID
      console.log(`⏳ 正在获取${name}的企信ID...`)
      const qixinResult = await getQixinId(id)

      if (qixinResult.returnCode === 1) {
        const qixinId = qixinResult.returnData
        console.log(`✅ 企信ID获取成功：${qixinId}`)

        // 调用接口发送PDF文件
        console.log(`⏳ 正在发送${pdfFile}到企信...`)
        const sendResult = await sendFileToUser(qixinId, fileFullPath, pdfFile)

        if (sendResult.code === 1) {
          console.log(`✅ 文件发送成功：${sendResult.message}`)
        } else {
          console.warn(`⚠️ 文件发送失败：${sendResult.message || "未知错误"}`)
        }
      } else {
        console.warn(`⚠️ 企信ID获取失败：${qixinResult.returnMessage}`)
      }

      successCount++
    } catch (error) {
      console.error(`❌ 处理失败 (${pdfFile}): ${error.message}`)
      failCount++
    }
  }

  console.log(`\n===== 提取统计 =====`)
  console.log(`✅ 成功：${successCount} 个`)
  console.log(`❌ 失败：${failCount} 个`)
}

// 直接执行函数
if (require.main === module) {
  sendWechatEeMsg().catch((error) => {
    console.error("❌ 提取失败：", error)
    process.exitCode = 1
  })
}

module.exports = { sendWechatEeMsg }