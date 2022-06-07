# IIS for connecting Exchange with Orchestrator

The script adds PSLanguageMode parameters to IIS via Powershell.

## Description

To correctly connect to Exchange servers from the Orchestrator, the following changes must be made:
* On the "Exchange Back End" site, change Powershell - "PSLanguageMode" from RestrictedLanguage to FullLanguage
* On the "Default Web Site" site, within the application, add the "PSLanguageMode" setting with the value "FullLanguage" in the Application Settings

## Getting Started

### Executing program

To run the program, just run the command in the proper path:

* if no file is attached
```
.\Set-IISForOrchestrator.ps1
```

* if file with config is attached
```
.\Set-IISForOrchestrator.ps1 .\EBEConfig_backup.csv
```



## Help

```
Get-Help Set-IISForOrchestrator.ps1
```

## Authors

[@MatekStatek](https://twitter.com/matekstatek)

## Version History

* 0.1
    * Initial Release