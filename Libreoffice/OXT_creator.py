from pathlib import Path
import zipfile

# Paths
output_oxt_path = Path("/mnt/data/CoolPropLibre_FromPy.oxt")
python_script_path = Path("/mnt/data/coolprop_libre.py")

# Create META-INF/manifest.xml
manifest_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="http://openoffice.org/2001/manifest">
    <manifest:file-entry manifest:media-type="application/vnd.sun.star.uno-component;type=Python" manifest:full-path="coolprop_libre.py"/>
</manifest:manifest>
"""

# Create description.xml
description_xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<description xmlns="http://openoffice.org/extensions/description/2006" xmlns:xlink="http://www.w3.org/1999/xlink">
    <identifier value="org.coolprop.libre"/>
    <version value="1.0.0"/>
    <display-name><name lang="en">CoolProp Libre</name></display-name>
    <extension-description><src lang="en" xlink:href="description.txt"/></extension-description>
</description>
"""

# Create the .oxt package
with zipfile.ZipFile(output_oxt_path, 'w') as oxt:
    oxt.writestr("META-INF/manifest.xml", manifest_xml)
    oxt.writestr("description.xml", description_xml)
    if python_script_path.exists():
        oxt.write(python_script_path, arcname="coolprop_libre.py")

output_oxt_path.name
