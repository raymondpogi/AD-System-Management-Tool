# Active Directory and System Management Tool

A PowerShell GUI tool for managing Active Directory users and groups, optimizing system performance, and troubleshooting Windows systems.

## Features
- **Active Directory Management**:
  - Retrieve detailed user information (`net user /domain`).
  - Search for users with partial name matches.
  - Get group details and search groups (`net group /domain`).
- **Device Optimization**:
  - Optimize power settings (disable hibernation, set timeouts to never).
  - Clean temporary files, Recycle Bin, Prefetch, Windows Update files, browser cache, and more.
  - Calculate and display freed disk space.
- **Troubleshooting**:
  - Start/stop services (`net start`, `net stop`).
  - Check BitLocker status and turn it off.
  - Uninstall Windows Updates by KB number.
  - Search Event Viewer logs (Application, Security, System) by Event ID.
  - Display recent Windows Update history.
- **System Information**:
  - Retrieve detailed computer information (brand, model, OS, memory, etc.).
- **File Management**:
  - Open file locations, delete files, and find the largest files in a folder.

## Requirements
- Windows 10 or later.
- PowerShell 5.1 or higher.
- Administrative privileges (required for most operations).
- Active Directory module (for AD queries).
- Winget (optional, for application updates).

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/AD-System-Management-Tool.git

![image](https://github.com/user-attachments/assets/102db8f6-eda4-40cf-b9a1-091d897055ab)
