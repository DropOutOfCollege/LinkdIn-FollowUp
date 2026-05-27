const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// ==================== 配置 ====================
const CONFIG = {
  DAILY_LIMIT: 200,
  MIN_DELAY: 3000,  // 3秒
  MAX_DELAY: 5000,  // 5秒
  COOKIES_PATH: './cookies.json',
  EXCEL_PATH: path.join(process.env.USERPROFILE || process.env.HOME, 'Desktop', `LinkedIn_FollowUp_${new Date().toISOString().split('T')[0]}.xlsx`),
  CHROME_PATH: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
};

// ==================== 工具函数 ====================

// 随机延迟（3-5秒）
async function randomDelay() {
  const delay = Math.floor(Math.random() * (CONFIG.MAX_DELAY - CONFIG.MIN_DELAY + 1)) + CONFIG.MIN_DELAY;
  await new Promise(resolve => setTimeout(resolve, delay));
}

// 偶尔长停顿（8-10秒）
async function occasionalLongDelay() {
  if (Math.random() < 0.1) { // 10% 概率
    const delay = Math.floor(Math.random() * 2000) + 8000;
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

// 根据语言匹配文档
function findDocument(docsPath, language) {
  if (!fs.existsSync(docsPath)) return null;
  
  const files = fs.readdirSync(docsPath);
  const patterns = {
    'en': [/en/i, /english/i, /_en/i],
    'es': [/es/i, /spanish/i, /español/i, /_es/i],
    'pt': [/pt/i, /portuguese/i, /português/i, /_pt/i],
    'zh': [/cn/i, /chinese/i, /zh/i, /_zh/i]
  };
  
  const patterns_for_lang = patterns[language] || patterns['en'];
  
  for (const file of files) {
    if (!file.endsWith('.pdf')) continue;
    
    for (const pattern of patterns_for_lang) {
      if (pattern.test(file)) {
        return path.join(docsPath, file);
      }
    }
  }
  
  // 如果没找到，返回第一个 PDF
  const firstPdf = files.find(f => f.endsWith('.pdf'));
  return firstPdf ? path.join(docsPath, firstPdf) : null;
}

// ==================== 核心逻辑 ====================

async function sendMessageToConnection(page, connection, messageTemplate, docsPath, excel, sentCount) {
  try {
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
    
    // 4. 输入消息
    console.log('  4. 输入消息...');
    const firstName = connection.name.split(' ')[0];
    const message = messageTemplate.replace(/{FirstName}/gi, firstName);
    
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
      
      // 点击附件按钮
      const attachBtn = await page.locator('button[aria-label*="attach"], button[aria-label*="file"]').first();
      if (attachBtn) {
        await attachBtn.click();
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
        if (sendBtn) break;
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
  
  // 初始化 Excel
  console.log('初始化 Excel...');
  const excel = initExcel();
  
  // 启动浏览器
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
    // 导航到 Connections 页面
    console.log('访问 Connections 页面...');
    await page.goto('https://www.linkedin.com/mynetwork/invite-connect/connections/');
    await randomDelay();
    await randomDelay();
    
    // 主循环
    while (sentCount < limit) {
      console.log(`\n--- 处理页面上的联系人 (${sentCount}/${limit}) ---`);
      
      // 获取当前页面的联系人
      const connections = await page.evaluate(() => {
        const buttons = document.querySelectorAll('button[aria-label^="More actions for"]');
        return Array.from(buttons).map(btn => {
          const ariaLabel = btn.getAttribute('aria-label');
          const name = ariaLabel.replace('More actions for ', '').trim();
          
          // 查找对应的链接
          let container = btn.parentElement;
          for (let i = 0; i < 8; i++) {
            if (!container) break;
            container = container.parentElement;
          }
          
          const link = container ? container.querySelector('a[href*="/in/"]') : null;
          const url = link ? link.href : '';
          
          return { name, url };
        }).filter(c => c.name && c.url);
      });
      
      console.log(`  页面上有 ${connections.length} 个联系人`);
      
      if (connections.length === 0) {
        console.log('  没有更多联系人，结束');
        break;
      }
      
      // 处理每个联系人
      for (const connection of connections) {
        if (sentCount >= limit) {
          console.log(`\n✅ 已达到每日上限 ${limit} 人，停止`);
          break;
        }
        
        processedCount++;
        
        // 发送消息
        const result = await sendMessageToConnection(page, connection, params.message, docsPath, excel, sentCount);
        
        if (result.success) {
          sentCount++;
        }
        
        // 返回 Connections 页面
        console.log('  返回 Connections 页面...');
        await page.goto('https://www.linkedin.com/mynetwork/invite-connect/connections/');
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
    
    // 关闭浏览器
    await browser.close();
    
    // 输出统计
    console.log('\n========================================');
    console.log('发送完成!');
    console.log('========================================');
    console.log(`处理联系人: ${processedCount}`);
    console.log(`成功发送: ${sentCount}`);
    console.log(`Excel 文件: ${CONFIG.EXCEL_PATH}`);
    console.log('========================================');
  }
}

// 运行主程序
main().catch(console.error);
