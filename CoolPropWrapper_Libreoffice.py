from CoolProp.CoolProp import PropsSI, HAPropsSI, get_global_param_string
from com.sun.star.sheet import XAddIn
import uno

def ConvertToSI(name, value):
    conversions = {
        "T": lambda x: x + 273.15,
        "P": lambda x: x * 1e5,
        "H": lambda x: x * 1000,
        "U": lambda x: x * 1000,
        "S": lambda x: x * 1000,
        "Cp": lambda x: x * 1000,
        "Cpmass": lambda x: x * 1000,
        "Cvmass": lambda x: x * 1000,
    }
    return conversions.get(name, lambda x: x)(value)

def ConvertFromSI(name, value):
    conversions = {
        "T": lambda x: x - 273.15,
        "P": lambda x: x / 1e5,
        "H": lambda x: x / 1000,
        "U": lambda x: x / 1000,
        "S": lambda x: x / 1000,
        "Cp": lambda x: x / 1000,
        "Cpmass": lambda x: x / 1000,
        "Cvmass": lambda x: x / 1000,
    }
    return conversions.get(name, lambda x: x)(value)

def FormatName(name):
    m = {
        "t": "T", "temp": "T", "temperature": "T",
        "p": "P", "pres": "P", "pressure": "P",
        "h": "H", "enth": "H", "enthalpy": "H", "hmass": "H",
        "u": "U", "internalenergy": "U", "umass": "U",
        "s": "S", "entr": "S", "entropy": "S", "smass": "S",
        "dmolar": "Dmolar", "dmol": "Dmolar", "delta": "Delta",
        "rho": "D", "dens": "D", "dmass": "D",
        "cvmass": "Cvmass", "cv": "Cvmass", "cp": "Cpmass", "cpmass": "Cpmass",
        "q": "Q", "quality": "Q", "x": "Q",
        "mu": "MU", "viscosity": "MU", "k": "K", "conductivity": "K",
        "z": "Z", "mm": "MM", "molar_mass": "MM", "phase": "Phase"
    }
    return m.get(name.lower(), name)

def FormatName_HA(name):
    m = {
        "twb": "Twb", "wetbulb": "Twb", "t_wb": "Twb",
        "tdp": "Tdp", "dewpoint": "Tdp", "t_dp": "Tdp",
        "t": "T", "tdb": "T", "t_db": "T",
        "p": "P", "p_w": "P_w", "r": "R", "rh": "R", "relhum": "R",
        "w": "W", "omega": "W", "humrat": "W",
        "hda": "Hda", "hha": "Hha", "sda": "Sda", "sha": "Sha",
        "cda": "Cda", "cpda": "Cda", "cha": "Cha", "cpha": "Cha",
    }
    return m.get(name.lower(), name)

def ConvertToSI_HA(name, value):
    if name in ("T", "Twb", "Tdp"):
        return value + 273.15
    if name in ("Hda", "Hha", "Sda", "Sha", "Cda", "Cha"):
        return value * 1000
    if name in ("P", "P_w"):
        return value * 1e5
    return value

def ConvertFromSI_HA(name, value):
    if name in ("T", "Twb", "Tdp"):
        return value - 273.15
    if name in ("Hda", "Hha", "Sda", "Sha", "Cda", "Cha"):
        return value / 1000
    if name in ("P", "P_w"):
        return value / 1e5
    return value

# === LibreOffice Calc Functions ===

def TMPr(output, name1, value1, name2, value2, fluid):
    try:
        if not all([output, name1, name2, fluid]):
            return "Error: Missing parameter"
        if not isinstance(value1, (int, float)) or not isinstance(value2, (int, float)):
            return "Error: Non-numeric inputs"

        name1 = FormatName(name1)
        name2 = FormatName(name2)
        output = FormatName(output)

        val1SI = ConvertToSI(name1, float(value1))
        val2SI = ConvertToSI(name2, float(value2))

        result = PropsSI(output, name1, val1SI, name2, val2SI, fluid)

        if abs(result) >= 1.0E+308 or result != result:
            return f"Error: CoolProp failed. {get_global_param_string('errstring')}"
        return ConvertFromSI(output, result)
    except Exception as e:
        return f"Error: {str(e)}"

def TMPa(output, name1, value1, name2, value2, name3, value3):
    try:
        if not all([output, name1, name2, name3]):
            return "Error: Missing parameter"
        if not all(isinstance(v, (int, float)) for v in [value1, value2, value3]):
            return "Error: Non-numeric inputs"

        name1 = FormatName_HA(name1)
        name2 = FormatName_HA(name2)
        name3 = FormatName_HA(name3)
        output = FormatName_HA(output)

        val1SI = ConvertToSI_HA(name1, float(value1))
        val2SI = ConvertToSI_HA(name2, float(value2))
        val3SI = ConvertToSI_HA(name3, float(value3))

        result = HAPropsSI(output, name1, val1SI, name2, val2SI, name3, val3SI)

        if abs(result) >= 1.0E+308 or result != result:
            return f"Error: CoolProp failed. {get_global_param_string('errstring')}"
        return ConvertFromSI_HA(output, result)
    except Exception as e:
        return f"Error: {str(e)}"
