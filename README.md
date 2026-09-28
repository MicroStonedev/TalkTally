# TalkTally

语音优先的个人记账工具——说一句话，就记好一笔账。

本仓库当前是 **MVP v1.0 的产品定义与可交互原型**，还没有可运行的 App 代码。

## 仓库结构

| 路径 | 内容 |
|---|---|
| [`docs/PRD.md`](docs/PRD.md) | MVP 产品文档：目标用户、痛点、核心功能、运作流程、技术方案要点、验证假设 |
| [`docs/prototype/`](docs/prototype/) | 可交互原型的界面截图与使用说明（原型本体在仓库根 `index.html`） |
| [`docs/poster/`](docs/poster/) | 产品手册海报（四联竖版）源文件、生成脚本与渲染图 |

## 原型试用

**在线（GitHub Pages）**：[https://microstonedev.github.io/TalkTally/](https://microstonedev.github.io/TalkTally/)

**本地**：用浏览器直接打开仓库根目录的 `index.html`，不需要服务器。演示路径见 [docs/prototype/README.md](docs/prototype/README.md)。

## 三条核心决策

1. **账本只存本地**：不强制账号，一键完整导出，防止产品自身变成下一个停服产品。
2. **智能不等于端侧**：解析分两层——高频句式走本地规则（零延迟、可复现、离线可用），长尾句才联网，且只发送当前这一句话，可一键关闭。
3. **无广告、无导流**：核心记账功能免费，导出不做成付费点。

## 状态

已完成：产品文档、交互原型、产品手册海报。
未开始：App 工程实现。分期计划、共享账本、多币种为 v0.2 范围。
