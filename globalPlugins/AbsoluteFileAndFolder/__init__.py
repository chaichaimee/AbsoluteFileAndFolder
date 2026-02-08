# __init__.py
# Copyright (C) 2026 'Chai Chaimee'
# Licensed under GNU General Public License. See COPYING.txt for details.

import time
import wx
import globalPluginHandler
import scriptHandler
import addonHandler
import core
import os
from . import AbsoluteFile
from . import AbsoluteFolder

# Initialize translation for the addon
addonHandler.initTranslation()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    # Category name for input gestures
    scriptCategory = _("Absolute File and Folder")

    def __init__(self):
        super().__init__()
        self._last_tap_time = 0
        self._tap_count = 0
        self._tap_threshold = 0.4  # seconds for double tap threshold
        
        # Check if should open last folders on startup
        core.callLater(3000, self._checkAndOpenLastFolders)

    def _checkAndOpenLastFolders(self):
        """Check and open folders that were opened before last shutdown only after Windows restart."""
        try:
            folder_manager = AbsoluteFolder.AbsoluteFolderManager()
            folder_manager.loadConfig()
            if folder_manager.shouldAutoOpenOnStartup() and folder_manager._lastOpenedFolders:
                # Open folders with a small delay between each
                for i, folder_path in enumerate(folder_manager._lastOpenedFolders):
                    if folder_path and os.path.isdir(folder_path):
                        core.callLater(1000 + (i * 500), lambda p=folder_path: self._openSingleFolder(p))
        except Exception:
            pass

    def _openSingleFolder(self, folder_path):
        """Open a single folder."""
        try:
            if folder_path and os.path.isdir(folder_path):
                os.startfile(folder_path)
        except Exception:
            pass

    @scriptHandler.script(
        description=_("Open Absolute Folders (single tap) or Absolute Files (double tap)"),
        category=_("Absolute File and Folder"),
        gesture="kb:windows+backspace"
    )
    def script_openAbsoluteManager(self, gesture):
        """
        Handle single tap for Folders and double tap for Files
        """
        current_time = time.time()
        
        # Reset count if time between taps is too long
        if current_time - self._last_tap_time > self._tap_threshold:
            self._tap_count = 0
        
        self._tap_count += 1
        self._last_tap_time = current_time
        
        def execute_action():
            if self._tap_count == 1:
                # Single tap: Folder Manager
                manager = AbsoluteFolder.AbsoluteFolderManager()
                manager.show()
            elif self._tap_count >= 2:
                # Double tap: File Manager
                manager = AbsoluteFile.AbsoluteFileManager()
                manager.show()
            # Reset after execution
            self._tap_count = 0

        # Wait for potential second tap before executing
        core.callLater(int(self._tap_threshold * 1000), execute_action)

    def terminate(self):
        super().terminate()

