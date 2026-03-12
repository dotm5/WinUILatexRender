# LocalLaTeXRender

一个基于 WinUI 3 和 WebView2 的本地 LaTeX 公式渲染工具，支持将公式转换为高 DPI 的透明 PNG，并写回系统剪贴板。

## 仓库结构

```text
.
|-- LocalLatexRender.sln
|-- README.md
|-- docs/
|-- src/
|   `-- LocalLatexRender/
|-- tests/
`-- artifacts/        # 本地构建、验证图片、诊断日志等，已被 Git 忽略
```

## 开发环境

- Windows 10/11
- .NET 10 SDK
- Windows App SDK / WinUI 3 开发环境

## 常用命令

```powershell
dotnet restore .\LocalLatexRender.sln
dotnet build .\LocalLatexRender.sln -p:Platform=x64
.\src\LocalLatexRender\bin\x64\Debug\net10.0-windows10.0.19041.0\win-x64\LocalLatexRender.exe
```

## GitHub 推送前建议

```powershell
git remote add origin <your-github-repo-url>
git status
git commit -m "Initial project structure"
git push -u origin main
```

更多目录说明见 [docs/repository-structure.md](./docs/repository-structure.md)。

