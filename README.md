# LinkedIn Follow-up Automation

## 安装

```bash
npm install
```

## 使用

### 1. 准备多语言文档

在指定目录放置文档，文件名包含语言标识：

```
D:\Follow-Up\Documents\
  ├── Cables_EN.pdf          (英语)
  ├── Cables_ES.pdf          (西班牙语)
  ├── Cables_PT.pdf          (葡萄牙语)
  └── Cables_CN.pdf          (中文)
```

### 2. 导出 LinkedIn Cookies

使用 Cookie Editor 插件导出 cookies.json 到脚本目录。

### 3. 运行

```bash
node linkedin-followup.js \
  --message "Hey {FirstName},\n\nLong time no contact..." \
  --docs-path "D:\\Follow-Up\\Documents" \
  --limit 200
```

### 参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--message` | 消息模板（含 `{FirstName}`） | 必填 |
| `--docs-path` | 多语言文档目录 | 必填 |
| `--limit` | 每日发送上限（最大200） | 200 |
| `--delay` | 基础延迟（秒） | 3-5随机 |
| `--excel-path` | Excel保存路径 | 桌面 |
| `--resume` | 从上次中断继续 | false |

## 输出

桌面生成 Excel 文件：`LinkedIn_FollowUp_YYYY-MM-DD.xlsx`

## 注意事项

- 需要提前登录 LinkedIn（使用 cookies.json）
- 每日上限 200 人
- 拟人化操作避免被封号
