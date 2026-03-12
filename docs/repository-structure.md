# Repository Structure

## 当前布局

- `src/LocalLatexRender/`
  - 主 WinUI 3 应用源码、资源、Web 渲染桥接文件。
- `docs/`
  - 面向维护者的项目说明、结构文档和后续设计文档。
- `tests/`
  - 预留给后续自动化测试、截图回归测试或集成测试。
- `artifacts/`
  - 本地产生的构建输出、日志和验证样图，默认不进入版本控制。

## 采用这个结构的原因

- `src` 与仓库根分离后，GitHub 上的代码入口更清晰。
- 解决方案文件位于根目录，方便 IDE、CI 和命令行统一入口。
- 文档、测试、产物目录职责单一，后续扩展更容易。

## 后续可扩展方向

- 在 `src/LocalLatexRender/Services/` 中拆分剪贴板、渲染和图像编码逻辑。
- 在 `src/LocalLatexRender/Models/` 中收纳渲染请求、历史记录、尺寸等数据模型。
- 在 `tests/` 中增加剪贴板 PNG 格式回归测试和 UI 自动化验证脚本。

