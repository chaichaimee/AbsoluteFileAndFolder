# __init__.py
# Copyright (C) 2026 Chai Chaimee
# Licensed under GNU General Public License. See COPYING.txt for details.

import time
import wx
import globalPluginHandler
import scriptHandler
import addonHandler
import core
import os
import shutil
import globalVars
import logHandler
from . import AbsoluteFile
from . import AbsoluteFolder

addonHandler.initTranslation()


def _migrate_config_files():
	config_path = globalVars.appArgs.configPath
	new_folder = os.path.join(config_path, "ChaiChaimee", "AbsoluteFileAndFloder")
	if not os.path.isdir(new_folder):
		try:
			os.makedirs(new_folder)
		except OSError as e:
			logHandler.log.warning(f"Failed to create config folder: {e}", exc_info=True)
			return
	old_files = [
		("AbsoluteFiles.json", "AbsoluteFiles.json"),
		("AbsoluteFolders.json", "AbsoluteFolders.json")
	]
	for old_name, new_name in old_files:
		old_path = os.path.join(config_path, old_name)
		new_path = os.path.join(new_folder, new_name)
		if os.path.isfile(old_path):
			try:
				shutil.move(old_path, new_path)
			except OSError as e:
				logHandler.log.warning(f"Failed to migrate {old_name}: {e}", exc_info=True)


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = _("Absolute File and Folder")

	def __init__(self):
		super().__init__()
		_migrate_config_files()
		self._pending_call_id = None
		self._last_tap_time = 0.0
		self._tap_count = 0
		self._tap_threshold = 0.4
		core.callLater(3000, self._checkAndOpenLastFolders)

	def _checkAndOpenLastFolders(self):
		try:
			folder_manager = AbsoluteFolder.AbsoluteFolderManager()
			folder_manager.loadConfig()
			if folder_manager.shouldAutoOpenOnStartup() and folder_manager._lastOpenedFolders:
				for i, folder_path in enumerate(folder_manager._lastOpenedFolders):
					if folder_path and os.path.isdir(folder_path):
						core.callLater(1000 + (i * 500), lambda p=folder_path: self._openSingleFolder(p))
		except Exception as e:
			logHandler.log.warning(f"Failed to check auto-open folders: {e}", exc_info=True)

	def _openSingleFolder(self, folder_path):
		try:
			if folder_path and os.path.isdir(folder_path):
				os.startfile(folder_path)
		except OSError as e:
			logHandler.log.warning(f"Failed to open folder on startup: {e}", exc_info=True)

	def _closeAnyOpenDialogs(self):
		folderDlg = AbsoluteFolder.AbsoluteFoldersDialog._activeInstance
		fileDlg = AbsoluteFile.AbsoluteFilesDialog._activeInstance
		if folderDlg and folderDlg.IsShown():
			folderDlg._silentClose()
		if fileDlg and fileDlg.IsShown():
			fileDlg._silentClose()

	@scriptHandler.script(
		description=_("Open Absolute Folders (single tap) or Absolute Files (double tap). If the target dialog is already open, brings it to the front instead of reopening it."),
		category=_("Absolute File and Folder"),
		gesture="kb:windows+backspace"
	)
	def script_openAbsoluteManager(self, gesture):
		current_time = time.time()
		if current_time - self._last_tap_time > self._tap_threshold:
			self._tap_count = 0
		self._tap_count += 1
		self._last_tap_time = current_time

		if self._pending_call_id is not None:
			self._pending_call_id.Cancel()
			self._pending_call_id = None

		def execute_action():
			self._pending_call_id = None
			if self._tap_count == 1:
				# Single tap targets the Folders dialog. If it is already
				# open, pull it to the foreground instead of closing and
				# recreating it.
				if AbsoluteFolder.AbsoluteFoldersDialog.bringToFront():
					self._tap_count = 0
					return
				self._closeAnyOpenDialogs()
				manager = AbsoluteFolder.AbsoluteFolderManager()
				manager.show()
			elif self._tap_count >= 2:
				# Double tap targets the Files dialog, same rule applies.
				if AbsoluteFile.AbsoluteFilesDialog.bringToFront():
					self._tap_count = 0
					return
				self._closeAnyOpenDialogs()
				manager = AbsoluteFile.AbsoluteFileManager()
				manager.show()
			self._tap_count = 0

		self._pending_call_id = core.callLater(int(self._tap_threshold * 1000), execute_action)

	def terminate(self):
		if self._pending_call_id is not None:
			self._pending_call_id.Cancel()
		super().terminate()