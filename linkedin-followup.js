const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// ==================== 配置 ====================
const CONFIG = {
  DAILY_LIMIT: 200,
  MIN_DELAY: 1500,  // 1.5秒（原3秒的一半）
  MAX_DELAY: 2500,  // 2.5秒（原5秒的一半）
  COOKIES_PATH: './cookies.json',
  EXCEL_PATH: path.join(process.env.USERPROFILE || process.env.HOME, 'Desktop', `LinkedIn_FollowUp_${new Date().toISOString().split('T')[0]}_${Date.now()}.xlsx`),
  CHROME_PATH: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
};

// ==================== 工具函数 ====================

// 随机延迟（3-5秒）
async function randomDelay() {
  const delay = Math.floor(Math.random() * (CONFIG.MAX_DELAY - CONFIG.MIN_DELAY + 1)) + CONFIG.MIN_DELAY;
  await new Promise(resolve => setTimeout(resolve, delay));
}

// 偶尔长停顿（4-5秒，原来是8-10秒）
async function occasionalLongDelay() {
  if (Math.random() < 0.1) { // 10% 概率
    const delay = Math.floor(Math.random() * 1000) + 4000;
    console.log(`  [拟人化] 长停顿 ${delay/1000} 秒...`);
    await new Promise(resolve => setTimeout(resolve, delay));
  }
}

// 加载 cookies
function loadCookies() {
  if (!fs.existsSync(CONFIG.COOKIES_PATH)) {
    throw new Error('未找到 cookies.json，请先导出 LinkedIn cookies');
  }
  
  let cookies = JSON.parse(fs.readFileSync(CONFIG.COOKIES_PATH, 'utf8'));
  
  // 修复 sameSite 值
  cookies = cookies.map(cookie => {
    if (cookie.sameSite === 'no_restriction') cookie.sameSite = 'None';
    else if (cookie.sameSite === 'unspecified') cookie.sameSite = 'Lax';
    else if (cookie.sameSite === 'lax') cookie.sameSite = 'Lax';
    delete cookie.storeId;
    return cookie;
  });
  
  return cookies;
}

// 初始化 Excel
function initExcel() {
  const headers = ['序号', '姓名', 'Profile URL', '语言', '消息内容', '附件名称', '发送时间', '状态', '备注'];
  
  const ws = xlsx.utils.aoa_to_sheet([headers]);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, '发送记录');
  
  return { wb, ws, rowCount: 1 };
}

// 追加记录到 Excel
function appendToExcel(excel, data) {
  const row = [
    excel.rowCount,
    data.name,
    data.url,
    data.language,
    data.message,
    data.attachment,
    data.time,
    data.status,
    data.note || ''
  ];
  
  xlsx.utils.sheet_add_aoa(excel.ws, [row], { origin: excel.rowCount });
  excel.rowCount++;
  
  // 实时保存
  xlsx.writeFile(excel.wb, CONFIG.EXCEL_PATH);
}

// 读取已发送记录（避免重复）
function loadSentHistory() {
  const history = new Set();
  
  // 查找桌面所有 LinkedIn_FollowUp Excel 文件
  const desktopPath = path.join(process.env.USERPROFILE || process.env.HOME, 'Desktop');
  
  if (fs.existsSync(desktopPath)) {
    const files = fs.readdirSync(desktopPath);
    const excelFiles = files.filter(f => f.startsWith('LinkedIn_FollowUp_') && f.endsWith('.xlsx'));
    
    for (const file of excelFiles) {
      try {
        const filePath = path.join(desktopPath, file);
        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.Sheets['发送记录'];
        
        if (sheet) {
          const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
          // 跳过表头，从第二行开始
          for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row && row[1]) { // 姓名在第2列
              history.add(row[1].toString().trim());
            }
          }
        }
      } catch (e) {
        console.log(`  读取历史记录失败: ${file}`);
      }
    }
  }
  
  console.log(`  已加载 ${history.size} 条历史发送记录`);
  return history;
}

// 检测语言
function detectLanguage(name, title) {
  const text = (name + ' ' + title).toLowerCase();
  
  // 西班牙语
  if (text.includes('de ') || text.includes('en ') || text.includes('logística') || 
      text.includes('licenciada') || text.includes('jefe') || text.includes('español')) {
    return 'es';
  }
  
  // 葡萄牙语
  if (text.includes('gerenciamento') || text.includes('especialista') || 
      text.includes('português') || text.includes('brazil')) {
    return 'pt';
  }
  
  // 中文
  if (/[\u4e00-\u9fa5]/.test(text)) {
    return 'zh';
  }
  
  // 默认英语
  return 'en';
}

// 根据语言匹配文档（支持 PDF 和 DOCX）
function findDocument(docsPath, language) {
  if (!fs.existsSync(docsPath)) return null;
  
  const files = fs.readdirSync(docsPath);
  const patterns = {
    'en': [/en/i, /english/i, /_en/i],
    'es': [/es/i, /spanish/i, /español/i, /_es/i],
    'pt': [/pt/i, /portuguese/i, /português/i, /_pt/i, /br/i],
    'zh': [/cn/i, /chinese/i, /zh/i, /_zh/i, /中文/i]
  };
  
  const patterns_for_lang = patterns[language] || patterns['en'];
  
  // 支持 PDF 和 DOCX
  const supportedExts = ['.pdf', '.docx'];
  
  for (const file of files) {
    const ext = path.extname(file).toLowerCase();
    if (!supportedExts.includes(ext)) continue;
    
    for (const pattern of patterns_for_lang) {
      if (pattern.test(file)) {
        return path.join(docsPath, file);
      }
    }
  }
  
  // 如果没找到，返回第一个支持的文件
  const firstFile = files.find(f => {
    const ext = path.extname(f).toLowerCase();
    return supportedExts.includes(ext);
  });
  return firstFile ? path.join(docsPath, firstFile) : null;
}

// ==================== 核心逻辑 ====================

async function sendMessageToConnection(page, connection, messageTemplate, docsPath, excel, sentCount, sentHistory) {
  try {
    // 检查是否已经发送过
    if (sentHistory.has(connection.name)) {
      console.log(`\n[${sentCount + 1}/${CONFIG.DAILY_LIMIT}] 跳过: ${connection.name} (已发送过)`);
      return { success: false, reason: 'already_sent' };
    }
    
    console.log(`\n[${sentCount + 1}/${CONFIG.DAILY_LIMIT}] 处理: ${connection.name}`);
    
    // 1. 导航到 Profile
    console.log('  1. 访问 Profile...');
    await page.goto(connection.url);
    await randomDelay();
    await occasionalLongDelay();
    
    // 2. 检测语言
    console.log('  2. 检测语言...');
    const lang = detectLanguage(connection.name, connection.title);
    console.log(`     检测到语言: ${lang.toUpperCase()}`);
    await randomDelay();
    
    // 3. 点击 Message
    console.log('  3. 点击 Message...');
    let messageBtn = null;
    
    // 尝试多个选择器
    const selectors = [
      'button:has-text("Message"):not(:has-text("InMail"))',
      'a:has-text("Message")',
      '[aria-label*="Message"]'
    ];
    
    for (const selector of selectors) {
      try {
        messageBtn = await page.waitForSelector(selector, { timeout: 3000 });
        if (messageBtn) break;
      } catch (e) {}
    }
    
    if (!messageBtn) {
      console.log('     ⚠️ 无法找到 Message 按钮，跳过');
      appendToExcel(excel, {
        name: connection.name,
        url: connection.url,
        language: lang,
        message: '',
        attachment: '',
        time: new Date().toLocaleString(),
        status: '跳过',
        note: '无法找到 Message 按钮'
      });
      return { success: false, reason: 'no_button' };
    }
    
    await messageBtn.click();
    await randomDelay();
    
    // 4. 输入消息 - 使用正确的名字
    console.log('  4. 输入消息...');
    // 从 connection.name 提取 first name
    const firstName = connection.name.split(' ')[0];
    // 替换模板中的 {FirstName}
    const message = messageTemplate.replace(/\{FirstName\}/gi, firstName);
    
    console.log(`     使用名字: ${firstName}`);
    
    let msgInput = null;
    const inputSelectors = [
      'div[contenteditable="true"]',
      'textarea',
      '.msg-form__contenteditable',
      '[role="textbox"]'
    ];
    
    for (const selector of inputSelectors) {
      try {
        msgInput = await page.waitForSelector(selector, { timeout: 3000 });
        if (msgInput) break;
      } catch (e) {}
    }
    
    if (!msgInput) {
      console.log('     ⚠️ 无法找到消息输入框');
      appendToExcel(excel, {
        name: connection.name,
        url: connection.url,
        language: lang,
        message: message,
        attachment: '',
        time: new Date().toLocaleString(),
        status: '失败',
        note: '无法找到输入框'
      });
      return { success: false, reason: 'no_input' };
    }
    
    await msgInput.fill(message);
    await randomDelay();
    
    // 5. 上传附件
    console.log('  5. 上传附件...');
    const docPath = findDocument(docsPath, lang);
    let attachmentName = '';
    
    if (docPath && fs.existsSync(docPath)) {
      attachmentName = path.basename(docPath);
      console.log(`     附件: ${attachmentName}`);
      
      // 点击附件按钮（使用 force 避免被拦截）
      const attachBtn = await page.locator('button[aria-label*="attach"], button[aria-label*="file"], button[title*="Attach"]').first();
      if (attachBtn) {
        try {
          await attachBtn.click({ force: true });
        } catch (e) {
          // 如果 force click 失败，尝试 JavaScript click
          await attachBtn.evaluate(el => el.click());
        }
        await randomDelay();
        
        // 上传文件
        const fileInput = await page.locator('input[type="file"]').first();
        if (fileInput) {
          await fileInput.setInputFiles(docPath);
          await randomDelay();
          await randomDelay(); // 额外等待上传完成
        }
      }
    } else {
      console.log('     ⚠️ 未找到对应语言文档');
    }
    
    // 6. 点击发送
    console.log('  6. 点击发送...');
    let sendBtn = null;
    const sendSelectors = [
      'button[type="submit"]',
      'button:has-text("Send")',
      '.msg-form__send-button',
      '.artdeco-button--primary'
    ];
    
    for (const selector of sendSelectors) {
      try {
        sendBtn = await page.waitForSelector(selector, { timeout: 3000 });
        if (sendBtn) {
          console.log(`     找到发送按钮: ${selector}`);
          break;
        }
      } catch (e) {}
    }
    
    if (!sendBtn) {
      console.log('     ⚠️ 无法找到发送按钮');
      appendToExcel(excel, {
        name: connection.name,
        url: connection.url,
        language: lang,
        message: message,
        attachment: attachmentName,
        time: new Date().toLocaleString(),
        status: '失败',
        note: '无法找到发送按钮'
      });
      return { success: false, reason: 'no_send' };
    }
    
    // 等待按钮可用
    await page.waitForTimeout(2000);
    
    // 尝试点击（最多3次）
    let sent = false;
    for (let i = 0; i < 3; i++) {
      try {
        const isEnabled = await sendBtn.evaluate(el => !el.disabled);
        if (isEnabled) {
          await sendBtn.click();
          sent = true;
          break;
        }
        await page.waitForTimeout(2000);
      } catch (e) {
        await page.waitForTimeout(2000);
      }
    }
    
    if (!sent) {
      // 强制点击
      try {
        await sendBtn.click({ force: true });
        sent = true;
      } catch (e) {}
    }
    
    if (sent) {
      console.log('     ✅ 发送成功');
      // 添加到历史记录
      sentHistory.add(connection.name);
      
      appendToExcel(excel, {
        name: connection.name,
        url: connection.url,
        language: lang,
        message: message,
        attachment: attachmentName,
        time: new Date().toLocaleString(),
        status: '成功',
        note: ''
      });
      await page.waitForTimeout(3000);
      return { success: true };
    } else {
      console.log('     ❌ 发送失败');
      appendToExcel(excel, {
        name: connection.name,
        url: connection.url,
        language: lang,
        message: message,
        attachment: attachmentName,
        time: new Date().toLocaleString(),
        status: '失败',
        note: '点击发送按钮失败'
      });
      return { success: false, reason: 'send_failed' };
    }
    
  } catch (error) {
    console.error(`     ❌ 错误: ${error.message}`);
    appendToExcel(excel, {
      name: connection.name,
      url: connection.url,
      language: 'unknown',
      message: '',
      attachment: '',
      time: new Date().toLocaleString(),
      status: '失败',
      note: error.message
    });
    return { success: false, reason: error.message };
  }
}

// ==================== 主程序 ====================

async function main() {
  // 解析参数
  const args = process.argv.slice(2);
  const params = {};
  
  for (let i = 0; i < args.length; i += 2) {
    const key = args[i].replace('--', '');
    const value = args[i + 1];
    params[key] = value;
  }
  
  // 验证必填参数
  if (!params.message) {
    console.error('错误: 请提供 --message 参数');
    console.log('示例: node linkedin-followup.js --message "Hey {FirstName}..." --docs-path "D:\\Documents"');
    process.exit(1);
  }
  
  const docsPath = params['docs-path'] || './documents';
  const limit = Math.min(parseInt(params.limit) || CONFIG.DAILY_LIMIT, CONFIG.DAILY_LIMIT);
  
  console.log('========================================');
  console.log('LinkedIn Follow-up Automation');
  console.log('========================================');
  console.log(`每日上限: ${limit} 人`);
  console.log(`文档路径: ${docsPath}`);
  console.log(`Excel 路径: ${CONFIG.EXCEL_PATH}`);
  console.log('========================================\n');
  
  // 加载 cookies
  console.log('加载 cookies...');
  const cookies = loadCookies();
  
  // 加载已发送历史
  console.log('加载已发送历史...');
  const sentHistory = loadSentHistory();
  
  // 初始化 Excel
  console.log('初始化 Excel...');
  const excel = initExcel();
  
  // 启动浏览器 - 只启动一次，复用同一个实例
  console.log('启动浏览器...');
  const browser = await chromium.launch({
    headless: false,
    executablePath: CONFIG.CHROME_PATH
  });
  
  const context = await browser.newContext({
    viewport: { width: 1280, height: 720 }
  });
  
  await context.addCookies(cookies);
  const page = await context.newPage();
  
  // 发送计数
  let sentCount = 0;
  let processedCount = 0;
  
  try {
    // 导航到 Connections 页面 - 只导航一次
    console.log('访问 Connections 页面...');
    await page.goto('https://www.linkedin.com/mynetwork/invite-connect/connections/');
    await randomDelay();
    await randomDelay();
    
    // 主循环
    while (sentCount < limit) {
      console.log(`\n--- 处理页面上的联系人 (${sentCount}/${limit}) ---`);
      
      // 获取当前页面的联系人 - 严格按照页面顺序
      const connections = await page.evaluate(() => {
        const results = [];
        
        // 查找所有连接卡片
        const cards = document.querySelectorAll('.mn-connection-card, li.artdeco-list__item');
        
        for (const card of cards) {
          // 获取姓名
          const nameEl = card.querySelector('.mn-connection-card__name, span[dir="ltr"]');
          const name = nameEl ? nameEl.textContent.trim() : '';
          
          // 获取职位
          const titleEl = card.querySelector('.mn-connection-card__occupation, .artdeco-entity-lockup__subtitle');
          const title = titleEl ? titleEl.textContent.trim() : '';
          
          // 获取链接
          const linkEl = card.querySelector('a[href*="/in/"]');
          const url = linkEl ? linkEl.href : '';
          
          // 只添加有效的联系人
          if (name && url && !url.includes('undefined')) {
            results.push({ name, title, url });
          }
        }
        
        return results;
      });
      
      console.log(`  页面上有 ${connections.length} 个联系人`);
      
      if (connections.length === 0) {
        console.log('  没有更多联系人，结束');
        break;
      }
      
      // 处理每个联系人 - 严格按照页面顺序
      for (const connection of connections) {
        if (sentCount >= limit) {
          console.log(`\n✅ 已达到每日上限 ${limit} 人，停止`);
          break;
        }
        
        processedCount++;
        
        // 发送消息
        const result = await sendMessageToConnection(page, connection, params.message, docsPath, excel, sentCount, sentHistory);
        
        if (result.success) {
          sentCount++;
        }
        
        // 返回 Connections 页面 - 使用浏览器的返回按钮
        console.log('  返回 Connections 页面...');
        await page.goBack();
        await randomDelay();
        
        // 额外随机停顿
        await occasionalLongDelay();
      }
      
      // 如果还没达到上限，滚动加载更多
      if (sentCount < limit) {
        console.log('\n--- 向下滚动加载更多 ---');
        await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
        await randomDelay();
        await randomDelay();
      }
    }
    
  } catch (error) {
    console.error('发生错误:', error.message);
  } finally {
    // 保存最终 Excel
    console.log('\n保存 Excel...');
    xlsx.writeFile(excel.wb, CONFIG.EXCEL_PATH);
    
    // 关闭浏览器 - 只关闭一次
    await browser.close();
    
    // 输出统计
    console.log('\n========================================');
    console.log('发送完成!');
    console.log('========================================');
    console.log(`处理联系人: ${processedCount}`);
    console.log(`成功发送: ${sentCount}`);
    console.log(`跳过(已发送): ${sentHistory.size}`);
    console.log(`Excel 文件: ${CONFIG.EXCEL_PATH}`);
    console.log('========================================');
  }
}

// 运行主程序
main().catch(console.error);
