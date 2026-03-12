# Contributing

Thanks for contributing to LocalLaTeXRender.

## Development Setup

1. Install .NET 10 SDK and the WinUI 3 / Windows App SDK toolchain.
2. Restore and build the solution:

```powershell
dotnet restore .\LocalLatexRender.sln
dotnet build .\LocalLatexRender.sln -p:Platform=x64
```

3. Run the app from the project output:

```powershell
.\src\LocalLatexRender\bin\x64\Debug\net10.0-windows10.0.19041.0\win-x64\LocalLatexRender.exe
```

## Pull Request Guidelines

- Keep changes focused and easy to review.
- Update documentation when behavior or project structure changes.
- Prefer small, descriptive commits.
- For UI or rendering changes, include before/after screenshots or PNG samples when possible.
- For clipboard or export changes, mention how you verified transparency, scaling, or compatibility.

## Code Style

- Follow the existing C# and XAML conventions in the repository.
- Prefer clear naming over clever abstractions.
- Avoid committing build outputs, local diagnostics, or generated artifacts.

