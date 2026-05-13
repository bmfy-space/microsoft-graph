# Microsoft Graph Email API

通过 Microsoft Graph API 获取 Outlook 邮件的 HTTP 服务，支持收件箱和垃圾邮件文件夹。

## 快速开始

```bash
pip install -r requirements.txt
python main.py
```

服务默认运行在 `http://localhost:3000`。

## API

### POST /api/email

获取邮件列表（支持分页，自动合并收件箱和垃圾邮件并去重）。

**请求参数（JSON Body）：**

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `credentials` | string | 是 | 格式：`email----password----clientId----refreshToken` |
| `page` | int | 否 | 页码，默认 `1` |
| `pageSize` | int | 否 | 每页数量，默认 `20` |
| `folders` | string[] | 否 | 邮件文件夹，默认 `["inbox", "junkemail"]` |

**调用示例：**

```bash
curl -X POST http://localhost:3000/api/email \
  -H "Content-Type: application/json" \
  -d '{
    "credentials": "user@outlook.com----password----clientId----refreshToken",
    "page": 1,
    "pageSize": 20
  }'
```

**返回示例：**

```json
{
  "code": 200,
  "msg": "success",
  "data": {
    "account": "user@outlook.com",
    "page": 1,
    "pageSize": 20,
    "total": 20,
    "emails": [
      {
        "id": "...",
        "internetMessageId": "...",
        "subject": "邮件主题",
        "from": { "name": "发件人", "address": "sender@example.com" },
        "to": ["收件人 <recipient@example.com>"],
        "cc": [],
        "date": "2026-05-13T10:30:00Z",
        "preview": "邮件预览文本...",
        "body": {
          "contentType": "html",
          "content": "<html>完整邮件内容</html>"
        }
      }
    ]
  }
}
```

**错误响应：**

| code | 说明 |
|------|------|
| 400 | 参数缺失或格式错误 |
| 401 | Token 刷新失败，凭证可能已失效 |
| 502 | Graph API 请求失败 |

## 依赖

- Python 3.10+
- Flask
- requests
