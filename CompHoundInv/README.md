# How to Register/Unregister

## Register

1. Build the project.

2. Copy add-in DLL file to one of following locations:

    - Anywhere, then *.addin file <Assembly> setting should be updated to the full path including the DLL filename.
    - Inventor `InstallPath\bin\` folder, then *.addin file `<Assembly>` setting should be the DLL name only: `AddInName.dll`.
    - Inventor `\InstallPath\bin\XX\` folder, then *.addin file `<Assembly>` setting shoule be a relative path: `XX\AddInName.dll`.

3. Copy the .addin manifest file to one of the following locations:

    - Inventor Version Dependent
        - Windows XP: `C:\Documents and Settings\All Users\Application Data\Autodesk\Inventor [version]\Addins\`
        - Windows7/Vista: `C:\ProgramData\Autodesk\Inventor [version]\Addins\`

    - Inventor Version Independent
        - Windows XP: `C:\Documents and Settings\All Users\Application Data\Autodesk\Inventor Addins\`
        - Windows7/Vista: `C:\ProgramData\Autodesk\Inventor Addins\`

    - Per User Override
        - Windows XP: `C:\Documents and Settings\<user>\Application Data\Autodesk\Inventor [version]\Addins\`
        - Windows7/Vista: `C:\Users\<user>\AppData\Roaming\Autodesk\Inventor [version]\Addins\`

4. Start up Inventor; the AddIn will automatically be found and loaded.

## Unregister

To unregister the AddIn, remove the Autodesk.<AddInName>.Inventor.addin file from the abovementioned .addin manifest file locations.
