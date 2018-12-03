# ExcelDNAWixSharp
Sample solution to create Wix# setup for ExcelDNA Addin with InstallScope (per user/machine setup).

- Build using VS 2017, [Paket] dependency manager, and [WixSharp] framework
- Based on CAs in [WixInstaller] and [ExcelDNAWixInstallerLM] as git submodules
- Configure wix\SetupInfo.wxi and Program.cs as needed
- Replace the res\*.ico,*.bmp, and *.rtf with your own assets
- It has only Install, Repair, Uninstall, and Upgrade (Upgrade is not tested yet)

[//]: # (These are reference links used in the body of this note and get stripped out when the markdown processor does its job. There is no need to format nicely because it shouldn't be seen. Thanks SO - http://stackoverflow.com/questions/4823468/store-comments-in-markdown-syntax)


[Paket]: <https://fsprojects.github.io/Paket/index.html>
[WixSharp]: <https://github.com/oleg-shilo/wixsharp>
[WixInstaller]: <https://github.com/Excel-DNA/WiXInstaller>
[ExcelDNAWixInstallerLM]: <https://github.com/bpatra/ExcelDNAWixInstallerLM>
