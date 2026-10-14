# CertForge · 批量证书生成工具

批量生成获奖证书的脚本工具集（Python + PowerShell），用于从结构化数据一键生成证书文档与图片。

## 脚本说明

| 脚本 | 作用 |
|---|---|
| `extract_certificate_data.py` | 从文档/表格中提取证书数据，整理为 JSON |
| `build_image_certificates.py` | 图片版证书批量生成 |
| `build_certificates.ps1` | PowerShell 版批量生成 |

## 数据格式

`certificate_data.json` 结构示例：

```json
{
  "issuer": "主办单位",
  "event": "赛事名称",
  "date": "2026年6月",
  "records": [
    {
      "award": "一等奖",
      "award_display": "一等奖",
      "work_title": "作品名称",
      "team_name": "队伍名",
      "members": "成员名单",
      "teacher": "指导教师"
    }
  ]
}
```

## 说明

- **隐私提示**：证书数据含真实姓名等个人信息，本仓库不提供原始数据文件，请按上述结构自备数据。
- 生成的成品（证书图片/文档）建议不要直接提交到公开仓库。

## 许可证

MIT © 2026 [ahdezneytiry-boop](https://github.com/ahdezneytiry-boop)
