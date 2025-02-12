using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

public static class CoolPropWrapper
{
    // Import CoolProp.dll
    [DllImport("CoolProp.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern double PropsSI(string output, string name1, double value1, string name2, double value2, string fluid);

    // Import CoolProp.dll
    [DllImport("CoolProp.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern double HAPropsSI(string output, string name1, double value1, string name2, double value2, string name3, double value3);

    // Function to calculate thermodynamic properties
    [ExcelFunction(Description = "Calculate thermodynamic properties of real fluids using CoolProp with engineering units.")]
    public static double TMPr(string output, string name1, double value1, string name2, double value2, string fluid)
    {
        // Convert inputs to SI units
        name1 = FormatName(name1); // Normalize name capitalization
        name2 = FormatName(name2); // Normalize name capitalization
        output = FormatName(output); // Normalize output name capitalization

        value1 = ConvertToSI(name1, value1);
        value2 = ConvertToSI(name2, value2);

        // Call CoolProp for the requested property
        double result = PropsSI(output, name1, value1, name2, value2, fluid);

        // Convert outputs to engineering units
        return ConvertFromSI(output, result);
    }

    // Function to calculate thermodynamic properties
    [ExcelFunction(Description = "Calculate thermodynamic properties of humid air using CoolProp with engineering units.")]
    public static double TMPa(string output, string name1, double value1, string name2, double value2, string name3, double value3)
    {
        // Convert inputs to SI units
        name1 = FormatName_HA(name1); // Normalize name capitalization
        name2 = FormatName_HA(name2); // Normalize name capitalization
        name3 = FormatName_HA(name3); // Normalize name capitalization
        output = FormatName_HA(output); // Normalize output name capitalization

        value1 = ConvertToSI_HA(name1, value1);
        value2 = ConvertToSI_HA(name2, value2);
        value3 = ConvertToSI_HA(name3, value3);

        // Call CoolProp for the requested property
        double result = HAPropsSI(output, name1, value1, name2, value2, name3, value3);

        // Convert outputs to engineering units
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
            case "Cvmass": // Mass specific heat at constant volume (kJ/kg/K to J/kg/K)
                return value * 1000;
            case "Dmolar": // Molar density (mol/m^3 to mol/m^3) - no conversion needed
                return value;
            case "D": // Mass density (kg/m^3 to kg/m^3) - no conversion needed
            case "Rho":
            case "Dmass":
                return value;
            case "Q": // Mass vapor quality (mol/mol to mol/mol) - no conversion needed
                return value;
            case "Tau": // Reciprocal reduced temperature (Tc/T) - no conversion needed
                return value;
            case "Umolar": // Molar specific internal energy (J/mol to J/mol) - no conversion needed
                return value;
            case "Smolar": // Molar specific entropy (J/mol/K to J/mol/K) - no conversion needed
                return value;
            case "Cpmolar": // Molar specific heat at constant pressure (J/mol/K to J/mol/K) - no conversion needed
                return value;
            case "Cp": // Mass specific heat at constant pressure (kJ/kg/K to J/kg/K)
            case "Cpmass":
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
            case "Cvmass": // Mass specific heat at constant volume (J/kg/K to kJ/kg/K)
                return value / 1000;
            case "Dmolar": // Molar density (mol/m^3 to mol/m^3) - no conversion needed
                return value;
            case "D": // Mass density (kg/m^3 to kg/m^3) - no conversion needed
            case "Rho":
            case "Dmass":
                return value;
            case "Q": // Mass vapor quality (mol/mol to mol/mol) - no conversion needed
                return value;
            case "Tau": // Reciprocal reduced temperature (Tc/T) - no conversion needed
                return value;
            case "Umolar": // Molar specific internal energy (J/mol to J/mol) - no conversion needed
                return value;
            case "Smolar": // Molar specific entropy (J/mol/K to J/mol/K) - no conversion needed
                return value;
            case "Cpmolar": // Molar specific heat at constant pressure (J/mol/K to J/mol/K) - no conversion needed
                return value;
            case "Cp": // Mass specific heat at constant pressure (J/kg/K to kJ/kg/K)
            case "Cpmass":
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
                return value + 273.15;
            case "Cda": // Specific heat of dry air (kJ/kg/K to J/kg/K)
                return value * 1000;
            case "Cha": // Specific heat of humid air (kJ/kg/K to J/kg/K)
                return value * 1000;
            case "Tdp": // Dew Point Temperature (°C to K)
                return value + 273.15;
            case "Hda": // Specific Enthalpy of dry air (kJ/kg to J/kg)
                return value * 1000;
            case "Hha": // Specific Enthalpy of humid air (kJ/kg to J/kg)
                return value * 1000;
            case "K": // Conductivity (W/m/K to W/m/K) - no conversion needed
                return value;
            case "MU": // Viscosity (mPa-s to Pa.s)
                return value / 1000;
            case "Psi_w": // Water molar fraction (mol water/mol humid air) - no conversion needed
                return value;
            case "P": // Pressure (bar to Pa)
                return value * 1e5;
            case "P_w": // Partial pressure of water (bar to Pa)
                return value * 1e5;
            case "R": // Relative Humidity (%) to fractional
                return value;
            case "Sda": // Specific entropy of dry air (kJ/kg/K to J/kg/K)
                return value * 1000;
            case "Sha": // Specific entropy of humid air (kJ/kg/K to J/kg/K)
                return value * 1000;
            case "T": // Dry Bulb Temperature (°C to K)
                return value + 273.15;
            case "Vda": // Specific volume of dry air (m3/kg to m3/kg) - no conversion needed
                return value;
            case "Vha": // Specific volume of humid air (m3/kg to m3/kg) - no conversion needed
                return value;
            case "W": // Humidity ratio (kg water/kg dry air) - no conversion needed
                return value;
            case "Z": // Compressibility factor - no conversion needed
                return value;
            case "Dda": // Density of dry air (kg/m3 to kg/m3) - no conversion needed
                return value;
            case "Dha": // Density of humid air (kg/m3 to kg/m3) - no conversion needed
                return value;
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
                return value - 273.15;
            case "Cda": // Specific heat of dry air (J/kg/K to kJ/kg/K)
                return value / 1000;
            case "Cha": // Specific heat of humid air (J/kg/K to kJ/kg/K)
                return value / 1000;
            case "Tdp": // Dew Point Temperature (K to °C)
                return value - 273.15;
            case "Hda": // Specific Enthalpy of dry air (J/kg to kJ/kg)
                return value / 1000;
            case "Hha": // Specific Enthalpy of humid air (J/kg to kJ/kg)
                return value / 1000;
            case "K": // Conductivity (W/m/K to W/m/K) - no conversion needed
                return value;
            case "MU": // Viscosity (Pa.s to mPa.s)
                return value * 1000;
            case "Psi_w": // Water molar fraction (mol water/mol humid air) - no conversion needed
                return value;
            case "P": // Pressure (Pa to bar)
                return value / 1e5;
            case "P_w": // Partial pressure of water (Pa to bar)
                return value / 1e5;
            case "R": // Relative Humidity (fractional to %)
                return value;
            case "Sda": // Specific entropy of dry air (J/kg/K to kJ/kg/K)
                return value / 1000;
            case "Sha": // Specific entropy of humid air (J/kg/K to kJ/kg/K)
                return value / 1000;
            case "T": // Dry Bulb Temperature (K to °C)
                return value - 273.15;
            case "Vda": // Specific volume of dry air (m3/kg to m3/kg) - no conversion needed
                return value;
            case "Vha": // Specific volume of humid air (m3/kg to m3/kg) - no conversion needed
                return value;
            case "W": // Humidity ratio (kg water/kg dry air) - no conversion needed
                return value;
            case "Z": // Compressibility factor - no conversion needed
                return value;
            case "Dda": // Density of dry air (kg/m3 to kg/m3) - no conversion needed
                return value;
            case "Dha": // Density of humid air (kg/m3 to kg/m3) - no conversion needed
                return value;
            default:
                return value; // No conversion if unit not found
        }
    }
    
}

