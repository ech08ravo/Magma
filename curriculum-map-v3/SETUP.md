# Curriculum Map Builder: one-time setup

This creates `Curriculum Map Builder.xlsm` in this folder (`curriculum-map-v3/`).
You do it once. After that, all code updates load from the `vba/` folder.

## Steps (Mac Excel)

1. **Allow code access.** Excel ▸ **Preferences** ▸ **Security** ▸ tick
   **Trust access to the VBA project object model**.
2. **Create the workbook.** File ▸ **New Blank Workbook**.
3. **Save it as macro-enabled.** File ▸ **Save As**:
   - Name: `Curriculum Map Builder`
   - Format: **Excel Macro-Enabled Workbook (.xlsm)**
   - Where: this folder, `GitHub/magma/curriculum-map-v3/`, next to the `vba` folder.
4. **Add the installer.** Tools ▸ **Macro** ▸ **Visual Basic Editor** (or Opt+F11).
   File ▸ **Import File…** ▸ choose `vba/modInstaller.bas`.
5. **Save** (Cmd+S) and close the Visual Basic Editor.

That's it. Tell Claude it's done; Claude loads the rest of the code and tests it.

If Excel asks to **grant access** to the `vba` folder, click **Grant Access** or
**Select**, then pick the `curriculum-map-v3` folder.

## Windows Excel (for colleagues later)

Same steps: File ▸ Options ▸ Trust Center ▸ Trust Center Settings ▸ Macro Settings ▸
tick **Trust access to the VBA project object model**; Alt+F11 opens the editor.
Colleagues only need this if they want to reload code; they don't need it to use the builder.

## Using the builder

Give colleagues **`Curriculum Map Builder (blank).xlsm`**: it has no course loaded, only a
start sheet with the **Import export…** button. They import their own dashboard export, then
Save As under their course name and keep the blank copy for next time.
`Curriculum Map Builder.xlsm` is the same builder with the S466 example loaded.
To remake the blank after a code change: open the example builder, run `ReloadModules`, save,
then run `MakeBlankBuilder` with the full path of the blank file.

- **Import export…** (button on Timeline and Tasks): choose a dashboard timeline export
  (.xlsx or .csv, from dashboard v7.4+ with unit guides loaded for guide weeks and GLOs).
  Planning columns you have already filled in (Tools, Assessment type, Planned Weeks, and any
  columns you added) are kept, matched by unit code + AT number.
- **Refresh views**: after editing Tasks or the Lists sheet, rebuilds every sheet (~20 s).
- **Click a GLO or mode label** on Timeline/Weeks to toggle that filter.
- The first time you import from a new folder, Mac Excel asks for file access: click Grant Access.

## Updating the code

Edit the files in `vba/`, then run `ReloadModules` (Tools ▸ Macro ▸ Macros…). Keep the .bas files
plain ASCII: Mac Excel imports them in a legacy encoding, so special characters are built with ChrW.
