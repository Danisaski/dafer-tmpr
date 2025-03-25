# CoolPropWrapper for Excel

## Installation
Move the following files wherever you want, if using [compiled binaries](https://github.com/Danisaski/CoolPropWrapper/tree/main/compiled):
```bash
.\compiled\{dotnet_framework}\CoolPropWrapper.xll
.\compiled\{dotnet_framework}\CoolPropWrapper.dll
.\CoolProp.dll
```
[Default release](https://github.com/Danisaski/CoolPropWrapper/releases/tag/alpha) corresponds to net6.0-windows.

Import in Excel the `.xll` file.

## Usage
CoolPropWrapper allows you to compute thermodynamic properties of real fluids and humid air using CoolProp in Excel. The primary functions available are:

### `TMPr(output, name1, value1, name2, value2, fluid)`
- Computes thermodynamic properties of real fluids.
- **Units used:**
  - Temperature: **Celsius (°C)** (converted internally to Kelvin (K))
  - Pressure: **bar** (converted internally to Pascal (Pa))
  - Enthalpy, Internal Energy, Entropy: **kJ/kg** (converted internally to J/kg)
  - Specific Heat: **kJ/kg/K** (converted internally to J/kg/K)

### Example Usage:
```excel
=TMPr("H", "T", 25, "P", 1.01325, "Water")
```
This will return the enthalpy (H) of water at **25°C** and **1.01325 bar**.

### `TMPa(output, name1, value1, name2, value2, name3, value3)`
- Computes thermodynamic properties of humid air.
- **Units used:**
  - Temperature: **Celsius (°C)** (converted internally to Kelvin (K))
  - Pressure: **bar** (converted internally to Pascal (Pa))
  - Enthalpy, Specific Heat, Entropy: **kJ/kg** (converted internally to J/kg)
  - Humidity Ratio: **kg_water/kg_dry_air**

### Example Usage:
```excel
=TMPa("W", "T", 25, "P", 1.01325, "RH", 0.5)
```
This will return the humidity ratio (W) for air at **25°C**, **1.01325 bar**, and **50% relative humidity**.

### Available Properties and Aliases

| **Property**            | **Aliases**                  | **Description**                                      | **Units in Add-in** | **SI/CoolProp Units** |
|------------------------|----------------------------|------------------------------------------------------|---------------------|--------------------|
| Temperature           | `T`, `temperature`         | Absolute temperature                                | °C                 | K                 |
| Pressure             | `P`, `pressure`            | Absolute pressure                                   | bar                | Pa                |
| Density              | `D`, `rho`, `dmass`        | Mass density                                       | kg/m³              | kg/m³             |
| Molar Density        | `Dmolar`                   | Molar density                                      | mol/m³             | mol/m³            |
| Enthalpy             | `H`, `hmass`, `enthalpy`   | Specific enthalpy                                  | kJ/kg              | J/kg              |
| Internal Energy      | `U`, `umass`               | Specific internal energy                           | kJ/kg              | J/kg              |
| Entropy              | `S`, `smass`               | Specific entropy                                   | kJ/kg·K            | J/kg·K            |
| Specific Heat (Cp)   | `Cpmass`, `Cp`            | Heat capacity at constant pressure                | kJ/kg·K            | J/kg·K            |
| Specific Heat (Cv)   | `Cvmass`, `Cv`            | Heat capacity at constant volume                  | kJ/kg·K            | J/kg·K            |
| Quality             | `Q`                         | Vapor quality (mass fraction of vapor)            | (unitless)         | (unitless)        |
| Speed of Sound       | `A`, `speed_of_sound`      | Speed of sound in the fluid                       | m/s                | m/s               |
| Thermal Conductivity | `K`, `conductivity`        | Thermal conductivity                              | W/m·K              | W/m·K             |
| Viscosity           | `MU`, `viscosity`          | Dynamic viscosity                                 | Pa·s               | Pa·s              |
| Prandtl Number      | `Prandtl`                  | Prandtl number                                    | (unitless)         | (unitless)        |
| Compressibility Factor | `Z`                     | Compressibility factor                            | (unitless)         | (unitless)        |
| Surface Tension     | `surface_tension`          | Surface tension                                   | N/m                | N/m               |
| Molar Mass          | `MM`, `molar_mass`         | Molar mass                                        | kg/kmol            | kg/kmol           |
| Critical Pressure   | `Pcrit`, `p_critical`      | Critical pressure                                 | bar                | Pa                |
| Critical Temperature| `Tcrit`                    | Critical temperature                              | °C                 | K                 |
| Triple Point Temp   | `Ttriple`                  | Triple point temperature                         | °C                 | K                 |
| Dew Point Temp      | `Tdp`, `dewpoint`          | Dew point temperature (humid air)                | °C                 | K                 |
| Wet Bulb Temp       | `Twb`, `wetbulb`           | Wet bulb temperature (humid air)                 | °C                 | K                 |
| Enthalpy (dry air)  | `Hda`                      | Specific enthalpy (dry air base)                 | kJ/kg              | J/kg              |
| Enthalpy (humid air)| `Hha`                      | Specific enthalpy (humid air)                    | kJ/kg            | J/kg            |
| Entropy (dry air)  | `Sda`                      | Specific entropy (dry air base)                   | kJ/kg·K              | J/kg·K              |
| Entropy (humid air)| `Sha`                      | Specific entropy (humid air)                     | kJ/kg·K            | J/kg·K            |
| Relative Humidity   | `RH`, `relhum`, `R`        | Relative humidity (humid air)                    | p.u.                  | p.u.                 |
| Humidity Ratio      | `W`, `omega`, `humrat`     | Humidity ratio (humid air)                       | kg_water/kg_dry_air | kg_water/kg_dry_air |



For more details on available properties and fluids, refer to the [CoolProp documentation](http://www.coolprop.org/).

## If building from source:
Restore packages from the `.csproj`
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

Import in Excel the `.xll` file.

