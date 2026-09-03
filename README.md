# Excel 职场减负工具箱

让职场人少做一点重复劳动。

这是一个面向 Windows 桌面版 Excel 的开源效率工具项目。当前首个功能是“文本框一键格式化”：选中文本框后，一次完成字号、粗体、字体颜色和填充颜色设置，并可保存个人偏好跨工作簿使用。

## 普通用户

1. 从 [Releases](https://github.com/speedpiao-wq/Excel-efficiency-toolkit/releases/latest) 下载最新 `.xlam` 加载项。
2. 阅读 [小白安装与使用指南](https://speedpiao-wq.github.io/Excel-efficiency-toolkit/)。
3. 在 Excel 的“加载项”选项卡中，将“一键格式化”固定到快速访问工具栏。

> 当前仓库正在准备首次公开发布。Release 下载入口会在 `v0.1.0` 发布后启用。

## 功能

- 全局可用，不绑定单个工作簿。
- 设置字号、是否加粗、字体颜色和文本框填充颜色。
- 支持一次选择多个文本框和组合中的文本框。
- 普通单元格或不支持的对象被选中时静默退出，不弹 VBA 异常。
- 设置按 Windows 用户保存，不写入业务工作簿。

## 社区与需求

- [稳定社区入口页](https://speedpiao-wq.github.io/Excel-efficiency-toolkit/community.html)：展示当前微信群二维码及备用反馈渠道。
- [提交功能需求](https://github.com/speedpiao-wq/Excel-efficiency-toolkit/issues/new?template=feature_request.yml)
- [报告问题](https://github.com/speedpiao-wq/Excel-efficiency-toolkit/issues/new?template=bug_report.yml)
- [查看 GitHub Discussions](https://github.com/speedpiao-wq/Excel-efficiency-toolkit/discussions)
- 建议启用 GitHub Discussions，作为永不过期、可检索的公开交流区。

微信群二维码只有短期有效期，因此项目不会把二维码直接写死在加载项或教程中。所有公开材料只链接到稳定社区入口页；维护者只需定期替换入口页中的二维码图片。

维护者更新二维码时运行：

```text
python scripts/update_wechat_qr.py "新的群二维码.png" --valid-until YYYY-MM-DD
```

定时工作流每周检查两次，在二维码未配置、即将过期或已经过期时创建一条维护 Issue。

## 给 Codex 用户

仓库内的 [`excel-global-addin-maker`](./skill/excel-global-addin-maker/) Skill 可以生成或修改 Excel 全局加载项，并要求在真实 Excel 中验证。

```text
$excel-global-addin-maker 帮我把重复的 Excel 操作制作成全局快捷按钮……
```

## 开发与验证

- VBA、Ribbon 与构建脚本位于 [`src`](./src/)。
- Skill 源码位于 [`skill`](./skill/)。
- Excel 运行测试位于 [`tests`](./tests/)。
- 发布步骤见 [`RELEASE_CHECKLIST.md`](./RELEASE_CHECKLIST.md)。

GitHub 托管运行器通常没有桌面版 Excel，因此自动化工作流只做仓库与 Skill 静态检查。`.xlam` 的最终发布必须在装有 Excel 的 Windows 环境中完成真实运行验证。

## 安全

`.xlam` 是可执行的 VBA 加载项。项目不会要求用户开启“启用所有宏”，也不会自动修改宏安全设置、受信任位置或 VBA 工程访问权限。公开分发边界见 [`SECURITY.md`](./SECURITY.md)。

## 许可

[MIT License](./LICENSE)

