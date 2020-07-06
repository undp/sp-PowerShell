# sp-PowerShell

[PowerShellRef]: https://docs.microsoft.com/en-us/powershell/
[PnPPowerShell]: https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online?view=sharepoint-ps

`sp-PowerShell` is a collection of scripts that perform various operations on SharePoint sites.

### Requirements

* [PowerShell 5][PowerShellRef]
* [PnP-PowerShell][PnPPowerShell] (Automatically imported by the sp-PowerShell module)

### Note

Does not currently support PowerShell 7 due to issues that pertain to PnP-PowerShell, which is being tracked [here](https://github.com/pnp/PnP-PowerShell/issues/2595).

### Installing

* Clone repository to corresponding [PowerShell Modules location](https://docs.microsoft.com/en-us/powershell/scripting/developer/module/installing-a-powershell-module).
* Run `Import-Module sp-PowerShell` and you are ready to go!

## Authors

* Daniel Tshin - **[@dantshin](https://github.com/dantshin)**
* Giuseppe Campanelli - **[@themilanfan](https://github.com/themilanfan)**

## Contributing

Please see the [Contribution Guide](CONTRIBUTING.md) for information on how to develop and contribute.

## License

sp-PowerShell is licensed under the [MIT License](LICENSE.md).