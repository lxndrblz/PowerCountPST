PowerCountPST is a PowerShell Script lets you count the number of elements found in an Outlook PST file. There are no limitations in terms of size and it's blazingly fast. Unlike other solutions, PowerCountPST searches recursively, so no matter how your files PST files are structured, it will always yield back the accurate amount. Once it's done it will unmount the PST file again, which makes it perfect to run in batches. I've successfully used this script in a production scenario with over 2000 PST files ranging from a couple of MB up to 90 GB.

## Requirements
This script depends on having Microsoft Outlook installed, as it's using Outlook's COM interface.

## Quick Start
1. Run the script in a PowerShell command prompt and provide a PST file as a parameter:
`.\PowerCountPsT.ps1 -pst C:\Temp\outlook.pst"`
2. Depending On the provided PST file, the output might look something like this:
```
Folder: \\test contains 3 items
Folder: \\test\Deleted Items contains 0 items
C:\Temp\outlook.pst
Total Items: 3
PSPath
------
Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Outlook.Application
0
3
```
## License

PowerCountPST is licensed under the [MIT license](https://github.com/lxndrblz/powercountpst/blob/master/LICENSE).

## Maintenance

This script is maintained by its author [Alexander Bilz](https://github.com/lxndrblz).

## Special Thanks

Go to Mitch [365 Guy](https://gallery.technet.microsoft.com/office/Count-number-of-emails-in-61a2748e), for his initial script to count the elements in a pst file. His script was the primer for writing a script that overcomes the size and file structure depth limitations.