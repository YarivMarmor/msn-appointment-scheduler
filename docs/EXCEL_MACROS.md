# Running the original Excel/VBA edition (unsigned macros)

The `legacy/Appointment_Scheduler_MSN_Tool.xlsm` file contains VBA macros and has **no trusted digital VBA publisher signature**. Download it only from a repository and maintainer you trust. Review the VBA sources in `legacy/` and scan the workbook before enabling anything.

## 1. Download and inspect

1. Open the workbook link in `README.md`, then click **Download raw file** in GitHub (or download the complete repository ZIP and extract it).
2. Keep the `.xlsm` extension. Open using **desktop Excel for Windows**, not Excel for the web (which cannot run VBA macros).
3. If Excel opens the file in **Protected View**, only select **Enable Editing** after verifying that it is the intended file. This does not necessarily enable macros.

## 2. If Excel shows a yellow 'Security Warning - Macros have been disabled'

If and only if you trust this exact workbook, click **Enable Content**. This makes this document trusted on supported Excel installations. Do **not** change global settings to 'Enable all macros'.

## 3. If Excel shows a red 'Security Risk - Microsoft has blocked macros ... source is untrusted'

Microsoft blocks many downloaded Internet files by default. For an individually reviewed, trusted workbook on Windows:

1. **Close Excel** and locate the `.xlsm` in Windows File Explorer.
2. Right-click the workbook -> **Properties** -> **General**.
3. If you see **Security: This file came from another computer...**, select **Unblock** -> **Apply** -> **OK**.
4. Reopen in Excel; if an **Enable Content** prompt now appears, enable it *only* if you trust this file.

The **Unblock** option may be absent (for example, depending on extraction method, file location, or organizational policy). Do not assume that means the file is safe. If your institution manages macro policy, contact its IT administrator.

## 4. Trusted locations (advanced; use sparingly)

In Excel: **File -> Options -> Trust Center -> Trust Center Settings -> Trusted Locations -> Add new location**. Use a dedicated local folder that contains **only reviewed files**. Excel may enable active content there without prompting, so avoid broad folders such as Downloads, Desktop, entire OneDrive folders, or shared drives. Organizations may prohibit this setting.

## 5. If the macros still do not run

- Confirm you are using desktop Excel on Windows and the `.xlsm` was not renamed `.xlsx`.
- Check for administrative policy that disables macros.
- In Excel, use **Alt+F11**, **Ctrl+R** to inspect `VBAProject`, Modules and the form. Do not modify the original workbook during troubleshooting.
- Open a fresh downloaded copy if your workbook was accidentally edited.

## Why we do not recommend 'Enable all macros'

Unsigned macros are executable code. Turning off protection for all Excel workbooks is dangerous. Neither a GitHub upload nor the presence of VBA source text proves safety. Always review what you run.

## Microsoft guidance

- [Macros from the Internet are blocked by default](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked)
- [Enable or disable macros in Microsoft 365 files](https://support.microsoft.com/en-us/office/vba/enable-or-disable-macros-in-microsoft-365-files)
- [Manage trusted locations](https://support.microsoft.com/en-us/office/security-privacy/add-remove-or-change-a-trusted-location-in-microsoft-office)
