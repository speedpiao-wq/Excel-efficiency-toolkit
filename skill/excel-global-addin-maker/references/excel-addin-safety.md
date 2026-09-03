# Excel add-in design and safety

## Compatibility boundary

- The bundled workflow targets Windows desktop Excel with VBA support.
- Excel for the web cannot execute VBA `.xlam` add-ins.
- A recipient needs macro execution permission, but does not need `AccessVBOM` merely to use a finished add-in.
- Programmatic building or editing of a VBA project may require Excel's “Trust access to the VBA project object model.” Ask the user to enable it temporarily if needed; never change that setting automatically.

## Reliable VBA patterns

- Work from `Selection`, `ActiveSheet`, or an explicitly supplied workbook rather than recorded object names.
- Use type checks before calling shape, range, chart, or table-specific APIs.
- Keep a narrow error boundary around each selected object so one protected or malformed object does not abort the remaining work.
- Use `TextFrame2.TextRange.Font` for modern Office text formatting and set the legacy `TextFrame.Characters.Font` too when compatibility requires it.
- For theme-stable custom colors, use explicit RGB values rather than theme color indexes.
- Preserve line, position, size, margins, and other properties unless the user asked to change them.

## Settings persistence

For small personal preferences, VBA `GetSetting` and `SaveSetting` provide a per-Windows-user store under the current user's profile. This keeps settings out of workbooks and lets the same add-in behave consistently across files. Defaults remain embedded in the add-in for first use and for sharing.

## Distribution and trust

- An `.xlam` is executable code. Explain what it changes and what it does not access.
- Keep source available for review when practical.
- Do not digitally sign with an untrusted or improvised certificate and present it as public trust.
- Files downloaded from email, chat, or the web may carry Windows' Internet-zone marker and be blocked. For a small trusted group, users can inspect the file and use Windows Properties → Unblock if appropriate. For broad distribution, use a reputable code-signing certificate and an installer or managed deployment.
- Never instruct users to enable all macros globally.

## Updating safely

Keep an unmodified backup, version filenames when behavior changes, close Excel before replacing an installed file, and re-run the same representative Excel checks after every change.
