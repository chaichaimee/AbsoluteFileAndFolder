<div align="center">

<img src="https://www.nvaccess.org/files/nvda/documentation/userGuide/images/nvda.ico" alt="NVDA Logo" width="120">

# Absolute File and Folder

Instant shortcut management, intelligent path detection, and automatic restart restoration for your essential files and folders in NVDA.

**author:** chai chaimee  
**url:** https://github.com/chaichaimee/AbsoluteFileAndFolder

</div>

## Introduction

**Absolute File and Folder** is a powerful productivity add-on for the NVDA screen reader designed to simplify access to your frequently used files and directories. By integrating directly with Windows Explorer, the add-on enables screen reader users to instantly capture selected file or folder paths, organize shortcuts, search history, launch executables with administrator privileges, and automatically restore active folders after a computer reboot.

### Hot Keys

> **Windows+Backspace**  
> Single Tap : Open Absolute Folders  
> Double Tap : Open Absolute Files

> **Smart Tap & Foreground Logic:**  
> • A single press within 400ms triggers the Folders manager, while pressing twice within 400ms triggers the Files manager.  
> • If the requested manager is already open on screen, pressing the hotkey brings the active window directly to the foreground and focuses the saved item list without recreating the window.  
> • Switching between Folders and Files automatically closes the active dialog before opening the selected one.

## Features

### 1. Intelligent Windows Explorer Context Capture

When you trigger the add-on while browsing in Windows Explorer, it automatically inspects the active shell window via COM automation to retrieve the path of the currently selected file or directory.

**How it works step-by-step:**

1. Navigate to any file or folder in Windows Explorer.
2. Press **Windows+Backspace** (Single Tap for Folders, Double Tap for Files).
3. Press the **Add** button in the dialog. The add-on pre-fills the prompt with the filename or folder name from Explorer, allowing you to save it instantly with custom display names.

### 2. Absolute Folders Management

Efficiently organize and access your primary folder directories with dedicated control options:

* **Saved Folders Tab:** Bookmark essential directories. Options include Add, Edit (rename), Remove, Pin/Unpin to top, and custom item reordering (Move Up / Move Down).
* **Recent Folders Tab:** Tracks up to 20 recently accessed folders for quick navigation. History can be cleared at any time.
* **Flexible Sorting & View Options:** Sort folders by Custom order, Ascending (a-z), or Descending (z-a). Check the *Show paths* option to view full directory paths alongside folder names.

### 3. Automatic Folder Restoration on System Restart

Never lose your working context after restarting your PC.

**Step-by-step logic:**

1. Enable the checkbox **Remember and open folders automatically on restart** in the Absolute Folders dialog.
2. When you open folders through the add-on, they are registered in the auto-open list.
3. The add-on monitors system uptime via Windows API (`GetTickCount64`). Upon detecting a system reboot, NVDA automatically reopens all remembered folders in Windows Explorer sequentially (staggered with a 1000ms initial delay + 500ms spacing) once NVDA starts up.
4. Manage or delete folders from the auto-open startup list directly via the list context menu or the **Delete** key.

### 4. Absolute Files Management & Category Filtering

Organize, search, and launch individual files with built-in category filters:

* **Filter Type Dropdown:** Quickly filter your saved or recent file lists by file type:
  * **All:** Display all saved files.
  * **Audio:** .mp3, .wav, .flac, .m4a, .ogg
  * **Video:** .mp4, .mkv, .avi, .mov
  * **Document:** .pdf, .docx, .txt, .xlsx, .pptx
  * **Code:** .py, .cpp, .java, .js, .html, .css
  * **Exe:** .exe, .bat, .cmd, .msi
* **Run as Administrator:** Right-click (or press Application key) on any executable or script file (.exe, .bat, .cmd, .msi) in the list and select *Run as Administrator* to launch it with elevated administrative privileges.

### 5. Accessible Search with Real-Time Audio Feedback

Type into the search field on either the Saved or Recent tabs to instantly filter results. The screen reader automatically announces matching results (e.g., *"3 matches found"* or *"12 files found"*) after a 500ms typing pause.

### 6. Auto-Dismiss Inactivity Timer

Dialogs stay on top for quick interaction but automatically close after 15 seconds of inactivity to keep your screen unburdened.

> **Note:** A gentle low beep plays when the dialog automatically times out. Any key press, mouse movement, or context menu interaction resets the 15-second timer.

### 7. Automatic Configuration Migration

Upgrading from previous versions is seamless. On startup, configuration files (`AbsoluteFiles.json` and `AbsoluteFolders.json`) are automatically migrated into a dedicated user folder (`%NVDA_CONFIG%/ChaiChaimee/AbsoluteFileAndFloder/`).

<div align="center">

## Support Me

If this tool has made your life easier, consider fueling the next update with a small donation.

[![Support me](https://img.shields.io/badge/Donate-Support%20Me-blue?style=for-the-badge&logo=stripe)](https://buy.stripe.com/dRm9AU1xQ3Ds22N6VK1VK01)

Your support means the world. Let's build something great together

&copy; 2026 Chai Chaimee NVDA Add-on Released under GNU GPL

</div>