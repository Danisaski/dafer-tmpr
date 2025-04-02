using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

public static class CoolPropWrapper
{
    // Import CoolProp.dll
    [DllImport("CoolProp.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern double PropsSI(string output, string name1, double value1, string name2, double value2, string fluid);

    // Import HAPropsSI
    [DllImport("CoolProp.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern double HAPropsSI(string output, string name1, double value1, string name2, double value2, string name3, double value3);

    // Import CoolProp's error retrieval function
    [DllImport("CoolProp.dll", CallingConvention = CallingConvention.Cdecl)]
    private static extern IntPtr get_global_param_string(string param);

    // Function to retrieve the error message from CoolProp
    private static string GetCoolPropError()
    {
        IntPtr ptr = get_global_param_string("errstring");
        return Marshal.PtrToStringAnsi(ptr) ?? "Unknown error";
    }

    // Function to calculate thermodynamic properties
    [ExcelFunction(Description = "Calculate thermodynamic properties of real fluids using CoolProp with engineering units.")]
    public static object TMPr(string output, string name1, object value1, string name2, object value2, string fluid)
    {
        // Check for missing or invalid inputs
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (value1 == null || !(value1 is double)) return "Error: First property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value2 == null || !(value2 is double)) return "Error: Second property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        // Normalize parameter names
        name1 = FormatName(name1);
        name2 = FormatName(name2);
        output = FormatName(output);

        // Convert inputs to SI units
        double val1SI = ConvertToSI(name1, (double)value1);
        double val2SI = ConvertToSI(name2, (double)value2);

        // Call CoolProp for the requested property
        double result;
        try
        {
            result = PropsSI(output, name1, val1SI, name2, val2SI, fluid);

            // Check if the result is a large number (potential error)
            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
            }
        }
        catch (Exception ex)
        {
            return $"Error: Exception occurred while computing property. {ex.Message}";
        }

        // Convert output to engineering units
        return ConvertFromSI(output, result);
    }


[ExcelFunction(Description = "Calculate thermodynamic properties of humid air using CoolProp with engineering units.")]
public static object TMPa(string output, string name1, object value1, string name2, object value2, string name3, object value3)
{
    // Check for missing or invalid inputs
    if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
    if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
    if (value1 == null || !(value1 is double)) return "Error: First property value is missing or not a number.";
    if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
    if (value2 == null || !(value2 is double)) return "Error: Second property value is missing or not a number.";
    if (string.IsNullOrWhiteSpace(name3)) return "Error: Third property name is missing.";
    if (value3 == null || !(value3 is double)) return "Error: Third property value is missing or not a number.";

    // Normalize parameter names
    name1 = FormatName_HA(name1);
    name2 = FormatName_HA(name2);
    name3 = FormatName_HA(name3);
    output = FormatName_HA(output);

    // Convert inputs to SI units
    double val1SI = ConvertToSI_HA(name1, (double)value1);
    double val2SI = ConvertToSI_HA(name2, (double)value2);
    double val3SI = ConvertToSI_HA(name3, (double)value3);

    // Call CoolProp for the requested property
    double result;
    try
    {
        result = HAPropsSI(output, name1, val1SI, name2, val2SI, name3, val3SI);

        // Check if the result is a large number (potential error)
        if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
        {
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
        }
    }
    catch (Exception ex)
    {
        return $"Error: Exception occurred while computing property. {ex.Message}";
    }

    // Convert output to engineering units
    return ConvertFromSI_HA(output, result);
    }


    // Normalize name capitalization to match the expected format
    private static string FormatName(string name)
    {
        switch (name.ToLower())
        {
            case "t":
            case "temperature":
                return "T";
            case "p":
            case "pressure":
                return "P";
            case "h":
            case "enthalpy":
            case "hmass":
                return "H";
            case "u":
            case "internalenergy":
            case "umass":
                return "U";
            case "s":
            case "entropy":
            case "smass":
                return "S";
            case "dmolar":
                return "Dmolar";
            case "delta":
                return "Delta";
            case "rho":
            case "dmass":
                return "D";
            case "cvmass":
                return "Cvmass";
            case "cpmass":
                return "Cpmass";
            case "cpmolar":
                return "Cpmolar";
            case "cvmolar":
                return "Cvmolar";
            case "q":
                return "Q";
            case "tau":
                return "Tau";
            case "alpha0":
                return "Alpha0";
            case "alphar":
                return "Alphar";
            case "speed_of_sound":
            case "a":
                return "A";
            case "bvirial":
                return "Bvirial";
            case "k":
            case "conductivity":
                return "K";
            case "cvirial":
                return "Cvirial";
            case "dipoe_moment":
                return "DIPOLE_MOMENT";
            case "fh":
                return "FH";
            case "g":
            case "gmass":
                return "G";
            case "helmoltzmass":
                return "HELMHOLTZMASS";
            case "helmholtzmolar":
                return "HELMHOLTZMOLAR";
            case "gamma":
                return "gamma";
            case "isobaric_expansion_coefficient":
                return "isobaric_expansion_coefficient";
            case "isothermal_compressibility":
                return "isothermal_compressibility";
            case "surface_tension":
                return "surface_tension";
            case "mm":
            case "molar_mass":
                return "MM";
            case "pcrit":
            case "p_critical":
                return "Pcrit";
            case "phase":
                return "Phase";
            case "pmax":
                return "pmax";
            case "pmin":
                return "pmin";
            case "prandtl":
                return "Prandtl";
            case "ptriple":
                return "ptriple";
            case "p_reducing":
                return "p_reducing";
            case "rhocrit":
                return "rhocrit";
            case "rhomass_reducing":
                return "rhomass_reducing";
            case "smolar_residual":
                return "Smolar_residual";
            case "tcrit":
                return "Tcrit";
            case "tmax":
                return "Tmax";
            case "tmin":
                return "Tmin";
            case "ttriple":
                return "Ttriple";
            case "t_freeze":
                return "T_freeze";
            case "t_reducing":
                return "T_reducing";
            case "mu":
            case "viscosity":
                return "MU";
            case "z":
                return "Z";
            default:
                return name; // Return as is if no match is found
        }
    }

    // Convert from custom units to SI units
    private static double ConvertToSI(string name, double value)
    {
        switch (name)
        {
            case "T": // Temperature (°C to K)
                return value + 273.15;
            case "P": // Pressure (bar to Pa)
                return value * 1e5;
            case "H": // Specific Enthalpy (kJ/kg to J/kg)
            case "U": // Specific Internal Energy (kJ/kg to J/kg)
            case "S": // Specific Entropy (kJ/kg to J/kg)
            case "Cp": // Mass specific heat at constant pressure (kJ/kg/K to J/kg/K)
            case "Cpmass":
            case "Cvmass": // Mass specific heat at constant volume (kJ/kg/K to J/kg/K)
                return value * 1000;
            default:
                return value; // No conversion if unit not found
        }
    }

    // Convert from SI units to custom units
    private static double ConvertFromSI(string name, double value)
    {
        switch (name)
        {
            case "T": // Temperature (K to °C)
                return value - 273.15;
            case "P": // Pressure (Pa to bar)
                return value / 1e5;
            case "H": // Specific Enthalpy (J/kg to kJ/kg)
            case "U": // Specific Internal Energy (J/kg to kJ/kg)
            case "S": // Specific Entropy (J/kg to kJ/kg)
            case "Cp": // Mass specific heat at constant pressure (J/kg/K to kJ/kg/K)
            case "Cpmass":
            case "Cvmass": // Mass specific heat at constant volume (J/kg/K to kJ/kg/K)
                return value / 1000;
            default:
                return value; // No conversion if unit not found
        }
    }

    // Normalize name capitalization for HA properties
    private static string FormatName_HA(string name)
    {
        switch (name.ToLower())
        {
            case "twb":
            case "t_wb":
            case "wetbulb":
                return "Twb";
            case "cda":
            case "cpda":
                return "Cda";
            case "cha":
            case "cpha":
                return "Cha";
            case "tdp":
            case "dewpoint":
            case "t_dp":
                return "Tdp";
            case "hda":
                return "Hda";
            case "hha":
                return "Hha";
            case "k":
            case "conductivity":
                return "K";
            case "mu":
            case "viscosity":
                return "MU";
            case "psi_w":
            case "y":
                return "Psi_w";
            case "p":
                return "P";
            case "p_w":
                return "P_w";
            case "r":
            case "rh":
            case "relhum":
                return "R";
            case "sda":
                return "Sda";
            case "sha":
                return "Sha";
            case "t":
            case "tdb":
            case "t_db":
                return "T";
            case "vda":
                return "Vda";
            case "vha":
                return "Vha";
            case "w":
            case "omega":
            case "humrat":
                return "W";
            case "z":
                return "Z";
            case "dda":
            case "rhoda":
                return "Dda";
            case "dha":
            case "rhoha":
                return "Dha";
            default:
                return name; // Return as is if no match is found
        }
    }

    // Convert from custom units to SI units for HA properties
    private static double ConvertToSI_HA(string name, double value)
    {
        switch (name)
        {
            case "Twb": // Wet Bulb Temperature (°C to K)
            case "Tdp": // Dew Point Temperature (°C to K)
            case "T": // Dry Bulb Temperature (°C to K)
                return value + 273.15;
            case "Cda": // Specific heat of dry air (kJ/kg/K to J/kg/K)
            case "Cha": // Specific heat of humid air (kJ/kg/K to J/kg/K)
            case "Hda": // Specific Enthalpy of dry air (kJ/kg to J/kg)
            case "Hha": // Specific Enthalpy of humid air (kJ/kg to J/kg)
            case "Sda": // Specific entropy of dry air (kJ/kg/K to J/kg/K)
            case "Sha": // Specific entropy of humid air (kJ/kg/K to J/kg/K)
                return value * 1000;
            case "P": // Pressure (bar to Pa)
            case "P_w": // Partial pressure of water (bar to Pa)
                return value * 1e5;
            default:
                return value; // No conversion if unit not found
        }
    }

    // Convert from SI units to custom units for HA properties
    private static double ConvertFromSI_HA(string name, double value)
    {
        switch (name)
        {
            case "Twb": // Wet Bulb Temperature (K to °C)
            case "Tdp": // Dew Point Temperature (K to °C)
            case "T": // Dry Bulb Temperature (K to °C)
                return value - 273.15;
            case "Cda": // Specific heat of dry air (J/kg/K to kJ/kg/K)
            case "Cha": // Specific heat of humid air (J/kg/K to kJ/kg/K)
            case "Hda": // Specific Enthalpy of dry air (J/kg to kJ/kg)
            case "Hha": // Specific Enthalpy of humid air (J/kg to kJ/kg)
            case "Sda": // Specific entropy of dry air (J/kg/K to kJ/kg/K)
            case "Sha": // Specific entropy of humid air (J/kg/K to kJ/kg/K)
                return value / 1000;
            case "P": // Pressure (Pa to bar)
            case "P_w": // Partial pressure of water (Pa to bar)
                return value / 1e5;
            default:
                return value; // No conversion if unit not found
        }
    }

}

