# Installation
Move the following files wherever you want, if using [compiled binaries](https://github.com/Danisaski/CoolPropWrapper/tree/main/compiled):
```bash
.\compiled\{dotnet_framework}\CoolPropWrapper.xll
.\compiled\{dotnet_framework}\CoolPropWrapper.dll
.\CoolProp.dll
```
[Default release](https://github.com/Danisaski/CoolPropWrapper/releases/tag/alpha) corresponds to net6.0-windows.

Import in excel the .xll file

## If building from source:
Restore packages from the .csproj
```bash
dotnet restore
```

Build release
```bash
dotnet build -c Release
```
Move the output files:
```bash
.\bin\Release\net6.0-windows\publish\CoolPropWrapper64-packed.xll
.\bin\Release\net6.0-windows\CoolPropWrapper.dll
.\CoolProp.dll
```

Import in excel the .xll file


