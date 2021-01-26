# Printers

Katy Nicholson, https://katystech.blog/

PowerShell script for mapping printers on logon with UI to show progress to end user. Runs multi-threaded to avoid UI lock up.
Runs the "Add Printer" step in a loop to retry on error (Windows 10 issue where printers randomly fail to map ~5% of the time with an unspecified error but if retried they work fine).

Printer mappings are set in a Group Policy Object which is not linked to any OU, using the Group Policy Preferences editor i.e. a graphical interface that technicians should be used to, and supports some item level targeting (Computer name, Group membership (computer or user), OU)

For detailed instructions see:
https://katystech.blog/2020/08/powershell-printer-script/
