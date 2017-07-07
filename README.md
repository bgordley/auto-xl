# AutoXL
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
