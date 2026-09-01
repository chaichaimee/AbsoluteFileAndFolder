# __init__.py
# Copyright (C) 2026 Chai Chaimee
# Licensed under GNU General Public License. See COPYING.txt for details.

import time
import threading
import wx
import globalPluginHandler
import scriptHandler
import addonHandler
import core
import os
import shutil
import json
import comtypes
import globalVars
import logHandler
from . import AbsoluteFile
from . import AbsoluteFolder

addonHandler.initTranslation()


def _migrateSingleFile(oldPath, newPath, layoutLabel, fileName):
	# Non-destructive by design (Section 13): copy to the new location first,
	# verify the copy is actually valid JSON, and only prune the old file once
	# the copy has been confirmed readable. A failed or partial write at the
	# destination therefore never costs the user their saved list, unlike a
	# straight shutil.move.
	try:
		shutil.copy2(oldPath, newPath)
	except OSError as e:
		logHandler.log.warning(f"Failed to copy {fileName} from {layoutLabel} layout during migration: {e}", exc_info=True)
		return
	try:
		with open(newPath, "r", encoding="utf-8") as f:
			json.load(f)
	except (OSError, ValueError) as e:
		logHandler.log.warning(f"Migrated copy of {fileName} from {layoutLabel} layout failed verification, leaving original in place: {e}", exc_info=True)
		return
	try:
		os.remove(oldPath)
	except OSError as e:
		logHandler.log.warning(f"Verified migration of {fileName} from {layoutLabel} layout, but failed to prune the original: {e}", exc_info=True)
		return
	logHandler.log.info(f"Migrated {fileName} from {layoutLabel} layout to the current config folder.")


def _migrate_config_files():
	config_path = globalVars.appArgs.configPath
	new_folder = os.path.join(config_path, "ChaiChaimee", "AbsoluteFileAndFolder")
	# Older builds stored this folder under a misspelled name; the affected
	# releases are treated as a distinct prior layout, checked before the
	# oldest flat-file layout, per the current-first / oldest-last order in
	# Section 13.
	misspelled_folder = os.path.join(config_path, "ChaiChaimee", "AbsoluteFileAndFloder")
	try:
		os.makedirs(new_folder, exist_ok=True)
	except OSError as e:
		logHandler.log.warning(f"Failed to create config folder: {e}", exc_info=True)
		return
	fileNames = ("AbsoluteFiles.json", "AbsoluteFolders.json")
	for fileName in fileNames:
		new_path = os.path.join(new_folder, fileName)
		if os.path.isfile(new_path):
			# Current layout already present for this file; migration for it
			# is a no-op so repeated loads stay idempotent.
			continue
		misspelled_path = os.path.join(misspelled_folder, fileName)
		if os.path.isfile(misspelled_path):
			_migrateSingleFile(misspelled_path, new_path, "misspelled-folder", fileName)
			continue
		legacy_flat_path = os.path.join(config_path, fileName)
		if os.path.isfile(legacy_flat_path):
			_migrateSingleFile(legacy_flat_path, new_path, "legacy flat-file", fileName)


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = _("Absolute File and Folder")

	def __init__(self):
		super().__init__()
		_migrate_config_files()
		self._pending_call_id = None
		self._startupCallId = None
		self._folderOpenCallIds = []
		self._last_tap_time = 0.0
		self._tap_count = 0
		self._tap_threshold = 0.4
		self._isTerminated = False
		self._startupCallId = core.callLater(3000, self._checkAndOpenLastFolders)

	def _checkAndOpenLastFolders(self):
		self._startupCallId = None
		if self._isTerminated:
			return
		try:
			folder_manager = AbsoluteFolder.AbsoluteFolderManager()
			folder_manager.loadConfig()
			if folder_manager.shouldAutoOpenOnStartup() and folder_manager._lastOpenedFolders:
				for i, folder_path in enumerate(folder_manager._lastOpenedFolders):
					if folder_path and os.path.isdir(folder_path):
						callId = core.callLater(1000 + (i * 500), self._openSingleFolder, folder_path)
						self._folderOpenCallIds.append(callId)
		except OSError as e:
			logHandler.log.warning(f"Failed to check auto-open folders: {e}", exc_info=True)

	def _openSingleFolder(self, folder_path):
		if self._isTerminated:
			return

		def worker():
			# Bracket os.startfile the same way the module-level shell helpers
			# bracket their COM access; ShellExecute (used internally by
			# os.startfile) can require an initialized apartment on the
			# calling thread for some file associations/shell verbs.
			comInitialized = False
			try:
				comtypes.CoInitialize()
				comInitialized = True
			except OSError:
				pass
			try:
				if folder_path and os.path.isdir(folder_path):
					os.startfile(folder_path)
			except OSError as e:
				logHandler.log.warning(f"Failed to open folder on startup: {e}", exc_info=True)
			finally:
				if comInitialized:
					try:
						comtypes.CoUninitialize()
					except OSError:
						pass
		threading.Thread(target=worker, daemon=True).start()

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
			self._pending_call_id.Stop()
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
		self._isTerminated = True
		if self._pending_call_id is not None:
			self._pending_call_id.Stop()
			self._pending_call_id = None
		if self._startupCallId is not None:
			self._startupCallId.Stop()
			self._startupCallId = None
		for callId in self._folderOpenCallIds:
			callId.Stop()
		self._folderOpenCallIds = []
		super().terminate()
