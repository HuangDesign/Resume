# 黄红艺 · 产品经理 · 简历

多年企业级产品经验，深耕 IoT、协同办公、数据处理等业务的数智化发展领域。近三年专注 AI Agent 与数据本体构建。

- 在线简历：<https://huangdesign.github.io/Resume/>
- 单文件版简历：[黄红艺-简历-产品经理.html](黄红艺-简历-产品经理.html)（双击即可打开，含内联资源，可离线发送）
- PDF 版：[黄红艺-简历-产品经理.pdf](黄红艺-简历-产品经理.pdf)
- GitHub：<https://github.com/HuangDesign>

## 文件说明

| 文件 | 说明 |
| --- | --- |
| `index.html` | 个人站点，GitHub Pages 首页 |
| `黄红艺-简历-产品经理.html` | A4 两页的简历，单文件、零外部依赖，可直接发邮件 |
| `黄红艺-简历-产品经理.pdf` | A4 两页打印版 |
| `黄红艺-简历-产品经理.docx` | Word 源文件 |
| `build_resume.py` | 用 python-docx 生成简历 DOCX |
| `assets/` | Logo 等静态资源 |

## 内容结构

- **01 个人优势**：对 AI 产品有经验 / 负责项目有结果 / 对 AI 产品设计有方法
- **02 工作经历**：锐捷网络、华为、涂鸦智能、海康威视
- **03 代表项目**（STAR）：华为 Answer 2.0 AI 办公、华为 Answer 1.0 AI 助手、锐捷 Source ID 3.0、华为 Talent 实训云平台
- **04 教育经历**、**05 专业技能**

## 重新生成

```bash
# 需要 python-docx
python3 build_resume.py                      # 生成 DOCX
soffice --headless --convert-to pdf 黄红艺-简历-产品经理.docx
```

PDF 与 DOCX 均为两页 A4（210 × 297 mm）。
