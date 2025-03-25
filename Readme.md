<!--toc:start-->
- [Restore packages from the .csproj (ExcelDna.Addin e.g.)](#restore-packages-from-the-csproj-exceldnaaddin-eg)
- [Build release](#build-release)
<!--toc:end-->


# Restore packages from the .csproj (ExcelDna.Addin e.g.)
```bash
dotnet restore
```

# Build release
```bash
dotnet build -c Release
```

# Installation
Move the output files wherever you want
```bash
.\bin\Release\net6.0-windows\publish\CoolPropWrapper64-packed.xll
.\bin\Release\net6.0-windows\CoolPropWrapper.dll
.\CoolProp.dll
```
Import in excel the .xll file
