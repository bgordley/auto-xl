# AutoXL
#### Summary:
AutoXL is a lightweight script that is designed to interact with Excel files without relying on heavy interop or other third-party dependencies. The core idea is that your Excel-specific automation happens inside of a VBA macro that is called by this script after basic error-checking.

#### Parameters:
- -c
  - Command to be executed
- -p
  - Path of settings.json file for command context

#### Commands: 
- INIT
  - Generates a settings.json file with placeholder elements
- RUN
  - Executes Excel automation using settings.json details
  
#### Usage:
```
powershell -command "& './AutoXL.ps1' INIT './path/to/settings.json'"
powershell -command "& './AutoXL.ps1' RUN './path/to/settings.json'"
```
