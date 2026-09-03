---
name: excel-global-addin-maker
description: Create, modify, package, and verify Windows desktop Excel .xlam add-ins for reusable global buttons, especially selection-based formatting and configurable text-box workflows. Use when a user wants an Excel action available across workbooks or wants to turn repetitive Excel steps into a Quick Access Toolbar button. Do not use for Excel for the web, which cannot run VBA add-ins.
---

# Excel Global Add-in Maker

Turn a repeated Excel desktop action into a portable `.xlam` add-in with a clear button, safe selection handling, and evidence from a real Excel verification.

## Choose the right output

- Use `.xlam` for Windows desktop Excel actions that must work across workbooks.
- Use `PERSONAL.XLSB` only when the user explicitly wants a private, single-machine macro rather than a shareable add-in.
- Explain that Excel for the web does not run VBA `.xlam` add-ins.
- If the user only needs to use an existing add-in, provide installation and Quick Access Toolbar guidance instead of rebuilding it.

## Build or modify

1. Extract the target selection, actions, configurable values, and properties that must remain unchanged. Ask only when ambiguity would materially change the result.
2. For the configurable text-box formatter, use the bundled builder:

   ```powershell
   & "<skill-folder>\scripts\build_textbox_formatter.ps1" `
     -OutputPath "<absolute-output-path>.xlam" `
     -DefaultFontSize 14 `
     -DefaultFontHex "#0033CC" `
     -DefaultFillHex "#FFFFFF" `
     -DefaultBold $true
   ```

3. For a different Excel action, copy the template to the task workspace and adapt the VBA and Ribbon callbacks. Keep the installed skill unchanged unless the user explicitly asks to update the skill itself.
4. For an existing `.xlam`, prefer editing available source. If only the binary exists, inspect its VBA project only when it is unlocked; never bypass a password or security protection.

Read [Excel add-in design and safety](references/excel-addin-safety.md) before changing VBA behavior, build permissions, or distribution assumptions. Read [beginner-friendly delivery](references/beginner-delivery.md) when the result will be shared with non-technical users.

## Required behavior

- Operate on the current user selection or active workbook; never hard-code workbook, worksheet, or shape names unless explicitly requested.
- Validate object types before formatting. A normal cell or unsupported object selection must be a silent no-op or a short controlled message, never an unhandled VBA error.
- Preserve unrelated formatting and workbook content.
- Store user preferences per Windows user when settings should persist without modifying workbooks.
- Put the main action and settings controls in a custom Ribbon group so users can right-click the main action and add it to the Quick Access Toolbar.
- Never change Excel macro security, trusted locations, registry security policy, or `AccessVBOM` automatically.

## Verify in real Excel

Use a new hidden Excel instance and a disposable unsaved workbook. At minimum:

1. Open the generated add-in.
2. Create a representative target object.
3. Apply non-default settings and run the public macro.
4. Read back every requested property from Excel.
5. Select a normal cell and confirm the macro exits without error.
6. Confirm the Ribbon XML, callbacks, and any settings form controls are present.
7. Restore the add-in's default per-user settings after testing.

Do not claim success from package inspection alone. If real Excel is unavailable, say that runtime verification remains incomplete.

## Deliver

Provide the `.xlam`, source or skill package when requested, and a novice guide covering installation, pinning, use, update, and uninstall. For public distribution, state that an unsigned Internet-downloaded add-in may be blocked and recommend code signing rather than weakening macro security.
