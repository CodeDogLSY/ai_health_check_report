const path = require("path")
const fs = require("fs-extra")
const iconv = require("iconv-lite")

/**
 * 读取 send_data 目录下的 PDF 文件，从文件名中提取姓名和证件号，
 * 调用外部接口获取企业微信ID，然后发送 PDF 文件到对应的企业微信账号。
 *
 * ✅ 关键修复（CentOS 7 / Linux 乱码文件名）：
 * Linux 文件名是“字节序列”，并不自带编码。
 * 如果真实文件名是 GBK/CP936 字节，而 Node 默认按 UTF-8 解码为 JS 字符串，
 * 会产生“�(U+FFFD)”替换字符，导致再用该字符串去访问文件时找不到文件。
 *
 * 本版本在 Linux 上使用 fs.readdir({ encoding: 'buffer' }) 保留原始字节，
 * 文件访问永远用原始字节（Buffer），展示/解析用解码后的 displayName。
 */
async function sendWechatEeMsg () {
  const ROOT = path.resolve(__dirname, ".")
  const SEND_DATA_DIR = path.join(ROOT, "send_data")
  if (!(await fs.pathExists(SEND_DATA_DIR))) {
    console.error(`❌ send_data 文件夹不存在`)
    return
  }

  console.log(`📋 正在读取send_data目录...`)
  console.log(`📋 send_data目录路径：${SEND_DATA_DIR}`)

  const isLinux = process.platform === "linux"
  if (isLinux)
    console.log(
      `🐧 检测到Linux系统，启用“原始字节文件名”模式（buffer readdir）`,
    )

  /** 判断解码结果是否像“体检报告_张三_身份证.pdf”这种结构 */
  function looksLikeReportFileName (s) {
    return (
      typeof s === "string" &&
      s.includes("报告") &&
      s.includes("_") &&
      /\d{15,18}/.test(s) &&
      s.toLowerCase().endsWith(".pdf")
    )
  }

  /**
   * Linux 下把“原始字节文件名”解码成用于展示/解析的中文名。
   * 注意：displayName 只是“展示名”，访问磁盘文件必须用 rawNameBuf。
   */
  function decodeDisplayNameFromRawBytes (rawNameBuf) {
    // 1) 常见编码优先尝试（你这个场景大概率是 CP936/GBK）
    const candidateEncodings = ["cp936", "gbk", "gb2312", "utf8", "big5"]
    for (const enc of candidateEncodings) {
      try {
        const decoded = iconv.decode(rawNameBuf, enc)
        if (looksLikeReportFileName(decoded)) return decoded
      } catch (_) {
        // ignore
      }
    }

    // 2) 兜底：先尝试 UTF-8，再用 latin1 显示（仅用于日志）
    try {
      const utf8 = rawNameBuf.toString("utf8")
      if (utf8 && utf8.toLowerCase().endsWith(".pdf")) return utf8
    } catch (_) { }
    return rawNameBuf.toString("latin1")
  }

  /**
   * 将目录路径（JS 字符串，UTF-8）+ 原始文件名（Buffer）拼成“原始字节完整路径”。
   * 这样 readFile/stat/pathExists 都能用正确的字节去访问文件。
   */
  function buildFullPathBuffer (dirPath, rawNameBuf) {
    // dirPath 基本都是 ASCII/UTF-8，Buffer.from 会按 UTF-8 编码
    return Buffer.concat([Buffer.from(dirPath + path.sep), rawNameBuf])
  }

  /**
   * 读取目录，返回统一结构：
   * - displayName: 用于解析姓名/证件号、作为上传 filename
   * - rawNameBuf/rawNameStr: 用于访问文件
   * - fullPath: 用于 fs.readFile/stat/pathExists（Linux: Buffer；Windows: string）
   */
  async function listPdfEntries () {
    if (isLinux) {
      const rawEntries = await fs.readdir(SEND_DATA_DIR, {
        encoding: "buffer",
      })
      const pdfEntries = []
      for (const rawNameBuf of rawEntries) {
        // 用 latin1/binary 做“扩展名过滤”最安全（只看 ASCII 的 .pdf）
        const rawLatin1 = rawNameBuf.toString("latin1")
        if (!rawLatin1.toLowerCase().endsWith(".pdf")) continue

        const displayName = decodeDisplayNameFromRawBytes(rawNameBuf)
        pdfEntries.push({
          displayName,
          rawNameBuf,
          rawLatin1,
          fullPath: buildFullPathBuffer(SEND_DATA_DIR, rawNameBuf),
        })
      }
      return pdfEntries
    }

    // Windows / macOS：文件名本身就是 JS 字符串（Node 已经给你正确 Unicode）
    const files = await fs.readdir(SEND_DATA_DIR)
    return files
      .filter((f) => path.extname(f).toLowerCase() === ".pdf")
      .map((f) => ({
        displayName: f,
        rawNameStr: f,
        fullPath: path.join(SEND_DATA_DIR, f),
      }))
  }

  const pdfEntries = await listPdfEntries()

  // 打印目录内容（帮助排障）
  if (isLinux) {
    // 注意：这里输出两份：rawLatin1(可见乱码) + displayName(解码后)
    console.log(
      `✅ 读取文件列表成功，共找到 ${pdfEntries.length} 个PDF文件（Linux: buffer）`,
    )
    console.log(
      `📋 文件系统中的实际文件名(raw latin1)：${JSON.stringify(pdfEntries.map((e) => e.rawLatin1))}`,
    )
    console.log(
      `📋 解码后的displayName：${JSON.stringify(pdfEntries.map((e) => e.displayName))}`,
    )
  } else {
    console.log(`✅ 读取文件列表成功，共找到 ${pdfEntries.length} 个PDF文件`)
    console.log(
      `📋 PDF文件列表：${JSON.stringify(pdfEntries.map((e) => e.displayName))}`,
    )
  }

  if (pdfEntries.length === 0) {
    console.log(`ℹ️ send_data 文件夹内没有PDF文件`)
    return
  }

  console.log(`📋 找到 ${pdfEntries.length} 个PDF文件，开始提取证件号...`)

  let successCount = 0
  let failCount = 0

  // 导入 axios 和 form-data 库
  const axios = require("axios")
  const FormData = require("form-data")

  // 获取企信ID
  async function getQixinId (sfz) {
    try {
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
      }
      throw new Error(`请求失败: ${error.message}`)
    }
  }

  // 上传并发送文件
  async function sendFileToUser (userId, fileFullPath, fileNameForUpload) {
    try {
      console.log(
        `📋 准备发送文件：${isLinux && Buffer.isBuffer(fileFullPath) ? fileFullPath.toString("latin1") : fileFullPath}`,
      )

      // 验证文件是否存在（fs-extra 通常支持 Buffer path；若不支持会抛错）
      let exists = false
      try {
        exists = await fs.pathExists(fileFullPath)
      } catch (_) {
        exists = false
      }
      if (!exists) {
        // 再用 stat 做一次兜底
        try {
          await fs.stat(fileFullPath)
          exists = true
        } catch (_) {
          exists = false
        }
      }
      if (!exists) {
        throw new Error(
          `文件不存在：${isLinux && Buffer.isBuffer(fileFullPath) ? fileFullPath.toString("latin1") : fileFullPath}`,
        )
      }

      const stats = await fs.stat(fileFullPath)
      if (!stats.isFile()) {
        throw new Error(
          `路径不是文件：${isLinux && Buffer.isBuffer(fileFullPath) ? fileFullPath.toString("latin1") : fileFullPath}`,
        )
      }
      console.log(`✅ 文件存在且可读取，大小：${stats.size} 字节`)

      const formData = new FormData()
      const fileBuffer = await fs.readFile(fileFullPath)
      formData.append("file", fileBuffer, {
        filename: fileNameForUpload,
        contentType: "application/pdf",
      })

      const response = await axios.post(
        `https://product.cajcare.com:5182/wechat/caj/sunflower/sendFileToUser?userId=${encodeURIComponent(userId)}`,
        formData,
        {
          headers: {
            ...formData.getHeaders(),
          },
          timeout: 30000,
        },
      )
      return response.data
    } catch (error) {
      if (error.response) {
        throw new Error(
          `发送文件失败: ${error.response.status} ${error.response.statusText}`,
        )
      } else if (error.request) {
        throw new Error(
          `发送文件失败: 没有收到响应，可能是网络超时或服务器问题`,
        )
      }
      throw new Error(`发送文件失败: ${error.message}`)
    }
  }

  // 主循环：逐个 PDF 处理
  for (let i = 0; i < pdfEntries.length; i++) {
    const entry = pdfEntries[i]
    const pdfFile = entry.displayName

    try {
      console.log(`📋 正在处理文件：${pdfFile}`)

      // 从文件名中提取姓名与证件号
      let name, id

      // 方法1：宽松正则
      const regex = /_([^_]+)_([\dXx]+)(?:_\d+)?\.pdf$/i
      const idMatch = pdfFile.match(regex)
      if (idMatch && idMatch.length >= 3) {
        name = idMatch[1]
        id = idMatch[2]
        console.log(`✅ 使用正则表达式提取成功：姓名=${name}，证件号=${id}`)
      } else {
        // 方法2：下划线分割
        console.log(`⚠️ 正则表达式匹配失败，尝试使用字符串分割法...`)
        const parts = pdfFile.replace(/\.pdf$/i, "").split("_")
        if (parts.length >= 3) {
          const idIndex = parts.findIndex((part) => /^[\dXx]+$/.test(part))
          if (idIndex > 0) {
            name = parts[idIndex - 1]
            id = parts[idIndex]
            console.log(
              `✅ 使用字符串分割法提取成功：姓名=${name}，证件号=${id}`,
            )
          } else {
            name = parts[parts.length - 2]
            id = parts[parts.length - 1]
            console.log(
              `⚠️ 无法确定证件号位置，尝试使用最后两部分：姓名=${name}，证件号=${id}`,
            )
          }
        } else {
          console.warn(`⚠️ 文件名格式不符合要求：${pdfFile}`)
          failCount++
          continue
        }
      }

      // 验证证件号格式
      if (!/^\d{15}$|^\d{17}[\dXx]$/.test(id)) {
        console.warn(`⚠️ 证件号格式不正确：${id}`)
        failCount++
        continue
      }

      // ✅ 关键：文件路径必须使用“原始字节路径”（Linux: Buffer；Windows: string）
      const fileFullPath = entry.fullPath
      console.log(
        `🔍 将发送的磁盘文件路径: ${isLinux && Buffer.isBuffer(fileFullPath) ? fileFullPath.toString("latin1") : fileFullPath}`,
      )
      if (isLinux && entry.rawLatin1) {
        console.log(`🔍 文件系统 raw 名称(latin1): ${entry.rawLatin1}`)
      }

      console.log(`✅ ${pdfFile} -> 姓名：${name}，证件号：${id}`)

      // 获取企信ID
      console.log(`⏳ 正在获取${name}的企信ID...`)
      const qixinResult = await getQixinId(id)
      if (qixinResult.returnCode !== 1) {
        console.warn(`⚠️ 企信ID获取失败：${qixinResult.returnMessage}`)
        failCount++
        continue
      }

      const qixinId = qixinResult.returnData
      console.log(`✅ 企信ID获取成功：${qixinId}`)

      // 发送文件：上传时的 filename 用解码后的中文名（pdfFile）
      console.log(`⏳ 正在发送${pdfFile}到企信...`)
      const sendResult = await sendFileToUser(qixinId, fileFullPath, pdfFile)
      if (sendResult.code === 1) {
        console.log(`✅ 文件发送成功：${sendResult.message}`)
      } else {
        console.warn(`⚠️ 文件发送失败：${sendResult.message || "未知错误"}`)
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
