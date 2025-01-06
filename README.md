# Create Day Folders by Year (PowerShell GUI)

A **PowerShell script** that creates **numbered month folders** (e.g., `1. January`, `2. February`, etc.) in **the same directory** as the script. It then generates subfolders (in `yyyy-MM-dd` format) for every date matching the **selected day of the week** in the **selected year**.

---

## Features

- **No folder selection required**  
  Folders are automatically created in the same directory that the script is in.
- **Numbered month folders**  
  Months are named `1. January`, `2. February`, ..., `12. December`.
- **GUI**  
  A small window allows you to pick:
  - **Year** (default: the current year)  
  - **Day of the Week** (default: Wednesday)  
- **Locale-safe day matching**  
  Uses `[System.DayOfWeek]` enums to avoid issues on non-English systems.

---

## Requirements

1. **Windows PowerShell** (v3 or later)  
   Included by default on Windows 7, 8, 10, 11, and most Windows Server editions.
2. **.NET Framework**  
   Must support Windows Forms (standard on most Windows installations).
3. **Permissions**  
   You may need to set `ExecutionPolicy` to allow running local scripts. For example:
   ```powershell
   Set-ExecutionPolicy Bypass -Scope Process
