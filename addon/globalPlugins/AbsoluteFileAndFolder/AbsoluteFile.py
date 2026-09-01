# AbsoluteFile.py

import os
import wx
import ui
import api
import gui
import globalVars
import json
import threading
import addonHandler
import core
import ctypes
from ctypes import wintypes
import comtypes
from comtypes import COMError as ComTypesCOMError
from comtypes.client import CreateObject as COMCreate
import urllib.parse
import logHandler

addonHandler.initTranslation()

TITLE = _("Absolute Files")

_HWND_TOPMOST = -1
_SWP_NOMOVE = 0x0002
_SWP_NOSIZE = 0x0001
_SWP_SHOWWINDOW = 0x0040


def _openFilePath(path):
	def worker():
		# ShellExecute (invoked internally by os.startfile) can require an
		# initialized COM apartment on the calling thread for some file
		# associations/shell verbs; bracket it the same way other Shell
		# COM access in this module is bracketed.
		comInitialized = False
		try:
			comtypes.CoInitialize()
			comInitialized = True
		except OSError:
			pass
		try:
			if path and os.path.isfile(path):
				os.startfile(path)
		except OSError as e:
			logHandler.log.warning(f"Failed to open file from Absolute Files: {e}", exc_info=True)
		finally:
			if comInitialized:
				try:
					comtypes.CoUninitialize()
				except OSError:
					pass
	threading.Thread(target=worker, daemon=True).start()


class AbsoluteFileManager:
	def __init__(self):
		self._files = {}
		self._order = []
		self._pinned = set()
		self._recentFiles = []
		self._showPath = False
		self._sortMode = "UPPERCASE"
		self._newFile = ""
		self.dialog = None
		self._is_resolving = False
		self._resolverThread = None
		self.loadConfig()

	def _get_config_path(self):
		folder = os.path.join(globalVars.appArgs.configPath, "ChaiChaimee", "AbsoluteFileAndFolder")
		return os.path.join(folder, "AbsoluteFiles.json")

	@staticmethod
	def _findFilePathInShellWindows(shell, windowHandle):
		for window in shell.Windows():
			try:
				if not window or window.hwnd != windowHandle:
					continue
				if hasattr(window, "Document") and window.Document:
					try:
						item = window.Document.FocusedItem
						if item:
							path = item.Path
							if path and os.path.isfile(path):
								return os.path.normpath(path)
					except (ComTypesCOMError, AttributeError):
						pass
				if hasattr(window, "LocationURL") and window.LocationURL:
					url = window.LocationURL
					if url.startswith("file:///"):
						path = urllib.parse.unquote(url[8:])
						path = path.replace("/", "\\")
						if os.path.isfile(path):
							return os.path.normpath(path)
			except (ComTypesCOMError, AttributeError, RuntimeError):
				continue
		return None

	def _getCurrentPathFromExplorer(self, fgAppName, fgHandle, focusAppName, focusHandle):
		comInitialized = False
		shell = None
		try:
			comtypes.CoInitialize()
			comInitialized = True
		except OSError:
			pass
		try:
			if fgAppName != "explorer" or not fgHandle:
				return None
			shell = COMCreate("Shell.Application")
			if not shell:
				return None
			path = self._findFilePathInShellWindows(shell, fgHandle)
			if path:
				return path
			if focusAppName == "explorer" and focusHandle:
				path = self._findFilePathInShellWindows(shell, focusHandle)
				if path:
					return path
		except (ComTypesCOMError, AttributeError, RuntimeError) as e:
			logHandler.log.warning(f"Failed to get Explorer file path: {e}", exc_info=True)
		finally:
			shell = None
			if comInitialized:
				try:
					comtypes.CoUninitialize()
				except OSError:
					pass
		return None

	def loadConfig(self):
		config_path = self._get_config_path()
		if os.path.isfile(config_path):
			try:
				with open(config_path, 'r', encoding='utf-8') as f:
					data = json.load(f)
				self._files = data.get("files", {})
				self._order = data.get("order", list(self._files.keys()))
				self._pinned = set(data.get("pinned", []))
				self._recentFiles = data.get("recentFiles", [])
				self._showPath = data.get("showPath", False)
				self._sortMode = data.get("sortMode", "UPPERCASE")
			except Exception as e:
				logHandler.log.warning(f"Failed to load file config: {e}", exc_info=True)

	def saveConfig(self):
		data = {
			"files": self._files,
			"order": self._order,
			"pinned": list(self._pinned),
			"recentFiles": self._recentFiles,
			"showPath": self._showPath,
			"sortMode": self._sortMode
		}
		config_path = self._get_config_path()
		try:
			os.makedirs(os.path.dirname(config_path), exist_ok=True)
			with open(config_path, 'w', encoding='utf-8') as f:
				json.dump(data, f, ensure_ascii=False, indent=2)
		except Exception as e:
			logHandler.log.error(f"Failed to save file config: {e}", exc_info=True)

	def addToRecent(self, path):
		if path and os.path.isfile(path):
			try:
				if path in self._recentFiles:
					self._recentFiles.remove(path)
				self._recentFiles.insert(0, path)
				self._recentFiles = self._recentFiles[:20]
				self.saveConfig()
			except Exception as e:
				logHandler.log.warning(f"Failed to add to recent files: {e}", exc_info=True)

	def show(self):
		if self._is_resolving:
			return
		self._is_resolving = True
		self.loadConfig()
		fg = api.getForegroundObject()
		fgAppName = fg.appModule.appName if fg and fg.appModule else None
		fgHandle = fg.windowHandle if fg else None
		focus = api.getFocusObject()
		focusAppName = focus.appModule.appName if focus and focus.appModule else None
		focusHandle = focus.windowHandle if focus else None

		def worker():
			try:
				path = self._getCurrentPathFromExplorer(fgAppName, fgHandle, focusAppName, focusHandle)
				wx.CallAfter(self._onPathResolved, path)
			finally:
				self._is_resolving = False
				self._resolverThread = None

		self._resolverThread = threading.Thread(target=worker, daemon=True)
		self._resolverThread.start()

	def _onPathResolved(self, path):
		if path and os.path.isfile(path):
			self._newFile = path
		else:
			self._newFile = ""
		self.dialog = AbsoluteFilesDialog(gui.mainFrame, self)
		gui.mainFrame.prePopup()
		self.dialog.CentreOnScreen()
		self.dialog.Show()
		self.dialog.Raise()
		self.dialog._applyTopMost()
		wx.CallAfter(self.dialog.listSaved.SetFocus)


class AbsoluteFilesDialog(wx.Dialog):
	_activeInstance = None

	def __init__(self, parent, manager):
		style = wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.STAY_ON_TOP
		super().__init__(parent, title=TITLE, style=style)
		self.timer = None
		if AbsoluteFilesDialog._activeInstance:
			AbsoluteFilesDialog._activeInstance._silentClose()
		AbsoluteFilesDialog._activeInstance = self
		self.manager = manager
		self._displayedRecentPaths = []
		self._contextMenuOpen = False
		self._pendingOpenPath = None
		self._filterGeneration = 0
		self._initUI()
		self._bindEvents()
		self.updateFiles()
		self.timer = wx.Timer(self)
		self.Bind(wx.EVT_TIMER, self.on_timeout, self.timer)
		self.Bind(wx.EVT_ACTIVATE, self.on_activate)
		self.timer.Start(15000)
		wx.CallAfter(self.listSaved.SetFocus)

	@classmethod
	def bringToFront(cls):
		inst = cls._activeInstance
		if inst is None:
			return False
		try:
			if not inst.IsShown():
				return False
		except RuntimeError:
			cls._activeInstance = None
			return False
		wx.CallAfter(inst._raiseToForeground)
		return True

	def _raiseToForeground(self):
		try:
			if self.IsIconized():
				self.Iconize(False)
			gui.mainFrame.prePopup()
			self.SetWindowStyleFlag(self.GetWindowStyleFlag() & ~wx.STAY_ON_TOP)
			self.SetWindowStyleFlag(self.GetWindowStyleFlag() | wx.STAY_ON_TOP)
			self.Raise()
			self.SetFocus()
			self._applyTopMost()
			self._reset_timer()
			if self.tabs.GetSelection() == 0:
				self.listSaved.SetFocus()
			else:
				self.listRecent.SetFocus()
		except (OSError, RuntimeError) as e:
			logHandler.log.warning(f"Failed to raise existing Files dialog: {e}", exc_info=True)

	def _applyTopMost(self):
		try:
			style = self.GetWindowStyleFlag()
			self.SetWindowStyleFlag(style & ~wx.STAY_ON_TOP)
			self.SetWindowStyleFlag(style | wx.STAY_ON_TOP)
			self.Raise()
		except (OSError, RuntimeError) as e:
			logHandler.log.warning(f"Failed to force Files dialog topmost: {e}", exc_info=True)

	def _play_close_beep(self):
		try:
			import winsound
			winsound.Beep(100, 100)
		except Exception:
			pass

	def _silentClose(self):
		self._destroyTimer()
		self._pendingOpenPath = None
		gui.mainFrame.postPopup()
		wx.CallAfter(self.Close)

	def _destroyTimer(self):
		if self.timer:
			try:
				self.timer.Stop()
			except Exception:
				pass
			try:
				self.Unbind(wx.EVT_TIMER, source=self.timer)
			except Exception:
				pass
			try:
				self.timer.Destroy()
			except Exception:
				pass
			self.timer = None

	def _reset_timer(self):
		if self.timer and not self._contextMenuOpen:
			self.timer.Stop()
			self.timer.Start(15000)

	def _stop_timer(self):
		if self.timer:
			self.timer.Stop()

	def on_activate(self, event):
		if event.GetActive():
			self._contextMenuOpen = False
			self._reset_timer()
			self._applyTopMost()
		event.Skip()

	def on_timeout(self, event):
		if not self._contextMenuOpen:
			self._play_close_beep()
			self.Close()

	def _initUI(self):
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		filterSizer = wx.BoxSizer(wx.HORIZONTAL)
		filterSizer.Add(wx.StaticText(self, label=_("Filter Type:")), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
		self.filterCombo = wx.ComboBox(self, choices=["All", "Audio", "Video", "Document", "Code", "Exe"], style=wx.CB_READONLY)
		self.filterCombo.SetSelection(0)
		filterSizer.Add(self.filterCombo, 1, wx.EXPAND)
		mainSizer.Add(filterSizer, 0, wx.EXPAND | wx.ALL, 5)

		self.tabs = wx.Notebook(self)
		self.panelSaved = wx.Panel(self.tabs)
		self.panelRecent = wx.Panel(self.tabs)
		self.tabs.AddPage(self.panelSaved, _("Saved Files"))
		self.tabs.AddPage(self.panelRecent, _("Recent Files"))
		mainSizer.Add(self.tabs, 1, wx.EXPAND | wx.ALL, 5)

		savedSizer = wx.BoxSizer(wx.VERTICAL)
		savedSearchSizer = wx.BoxSizer(wx.HORIZONTAL)
		savedSearchSizer.Add(wx.StaticText(self.panelSaved, label=_("Search:")), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
		self.searchSaved = wx.TextCtrl(self.panelSaved)
		savedSearchSizer.Add(self.searchSaved, 1, wx.EXPAND)
		savedSizer.Add(savedSearchSizer, 0, wx.EXPAND | wx.ALL, 5)
		self.listSaved = wx.ListCtrl(self.panelSaved, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.BORDER_SUNKEN)
		self.listSaved.InsertColumn(0, _("Name"), width=250)
		self.listSaved.InsertColumn(1, _("Path"), width=400)
		savedSizer.Add(self.listSaved, 1, wx.EXPAND | wx.ALL, 5)

		savedBtnSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnAdd = wx.Button(self.panelSaved, label=_("&Add"))
		self.btnEdit = wx.Button(self.panelSaved, label=_("&Edit"))
		self.btnRemove = wx.Button(self.panelSaved, label=_("&Remove"))
		savedBtnSizer.Add(self.btnAdd, 0, wx.RIGHT, 5)
		savedBtnSizer.Add(self.btnEdit, 0, wx.RIGHT, 5)
		savedBtnSizer.Add(self.btnRemove, 0)
		savedSizer.Add(savedBtnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 5)
		self.panelSaved.SetSizer(savedSizer)

		recentSizer = wx.BoxSizer(wx.VERTICAL)
		recentSearchSizer = wx.BoxSizer(wx.HORIZONTAL)
		recentSearchSizer.Add(wx.StaticText(self.panelRecent, label=_("Search:")), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
		self.searchRecent = wx.TextCtrl(self.panelRecent)
		recentSearchSizer.Add(self.searchRecent, 1, wx.EXPAND)
		recentSizer.Add(recentSearchSizer, 0, wx.EXPAND | wx.ALL, 5)
		self.listRecent = wx.ListCtrl(self.panelRecent, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.BORDER_SUNKEN)
		self.listRecent.InsertColumn(0, _("File Name"), width=250)
		self.listRecent.InsertColumn(1, _("Path"), width=400)
		recentSizer.Add(self.listRecent, 1, wx.EXPAND | wx.ALL, 5)
		self.btnClearRecent = wx.Button(self.panelRecent, label=_("Clear History"))
		recentSizer.Add(self.btnClearRecent, 0, wx.ALIGN_RIGHT | wx.ALL, 5)
		self.panelRecent.SetSizer(recentSizer)

		optionsSizer = wx.BoxSizer(wx.HORIZONTAL)
		choices = [_("Custom order"), _("Ascending, a-z"), _("Descending z-a")]
		self.sortCombo = wx.ComboBox(self, choices=choices, style=wx.CB_READONLY)
		mode_map = {"CUSTOM": 0, "UPPERCASE": 1, "LOWERCASE": 2}
		self.sortCombo.SetSelection(mode_map.get(self.manager._sortMode, 1))
		optionsSizer.Add(self.sortCombo, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
		mainSizer.Add(optionsSizer, 0, wx.EXPAND | wx.ALL, 5)

		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.btnOpen = wx.Button(self, label=_("&Open"))
		self.btnClose = wx.Button(self, wx.ID_CLOSE)
		btnSizer.Add(self.btnOpen)
		btnSizer.Add(self.btnClose)
		mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

		self.SetSizer(mainSizer)
		self.SetMinSize((800, 600))
		self.Fit()

	def _bindEvents(self):
		self.filterCombo.Bind(wx.EVT_COMBOBOX, lambda e: self.updateFiles() or self._reset_timer())
		self.searchSaved.Bind(wx.EVT_TEXT, self.onSearchTextChanged)
		self.searchRecent.Bind(wx.EVT_TEXT, self.onSearchTextChanged)
		self.sortCombo.Bind(wx.EVT_COMBOBOX, self.onSortChanged)
		self.tabs.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.onTabChanged)
		self.btnOpen.Bind(wx.EVT_BUTTON, self.onOpen)
		self.btnClose.Bind(wx.EVT_BUTTON, lambda e: self.Close())
		self.btnAdd.Bind(wx.EVT_BUTTON, self.onAdd)
		self.btnEdit.Bind(wx.EVT_BUTTON, self.onEdit)
		self.btnRemove.Bind(wx.EVT_BUTTON, self.onRemove)
		self.btnClearRecent.Bind(wx.EVT_BUTTON, self.onClearRecent)
		self.listSaved.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.onOpen)
		self.listRecent.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.onOpen)
		self.listSaved.Bind(wx.EVT_CONTEXT_MENU, self.onContextMenu)
		self.listRecent.Bind(wx.EVT_CONTEXT_MENU, self.onContextMenu)
		self.listSaved.Bind(wx.EVT_KEY_DOWN, self.onKeyDown)
		self.listRecent.Bind(wx.EVT_KEY_DOWN, self.onKeyDown)
		self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)
		self.Bind(wx.EVT_CLOSE, self.on_close)

	def on_close(self, event):
		self._destroyTimer()
		if AbsoluteFilesDialog._activeInstance == self:
			AbsoluteFilesDialog._activeInstance = None
		pendingPath = self._pendingOpenPath
		self._pendingOpenPath = None
		event.Skip()
		if pendingPath and os.path.isfile(pendingPath):
			core.callLater(300, lambda p=pendingPath: _openFilePath(p))

	def onTabChanged(self, evt):
		self.updateFiles()
		if self.tabs.GetSelection() == 0:
			self.listSaved.SetFocus()
		else:
			self.listRecent.SetFocus()
		self._reset_timer()

	def onSortChanged(self, evt):
		idx = self.sortCombo.GetSelection()
		self.manager._sortMode = ["CUSTOM", "UPPERCASE", "LOWERCASE"][idx]
		self.manager.saveConfig()
		self.updateFiles()
		self._reset_timer()

	def onCharHook(self, evt):
		if evt.GetKeyCode() == wx.WXK_ESCAPE:
			self._stop_timer()
			self.Close()
			return
		else:
			evt.Skip()
		self._reset_timer()

	def onSearchTextChanged(self, evt):
		self._reset_timer()
		self._filterGeneration += 1
		currentGeneration = self._filterGeneration
		currentText = self.searchSaved.GetValue() if self.tabs.GetSelection() == 0 else self.searchRecent.GetValue()
		self.updateFiles()
		core.callLater(500, self._announceSearchResultIfCurrent, currentGeneration, currentText)
		evt.Skip()

	def _announceSearchResultIfCurrent(self, generation, filterText):
		if generation != self._filterGeneration:
			return
		try:
			if self.tabs.GetSelection() == 0:
				count = self.listSaved.GetItemCount()
			else:
				count = self.listRecent.GetItemCount()

			if filterText:
				ui.message(_("{0} matches found").format(count))
			else:
				ui.message(_("{0} files found").format(count))
		except RuntimeError:
			pass

	def onKeyDown(self, evt):
		if evt.GetKeyCode() == wx.WXK_DELETE:
			if self.tabs.GetSelection() == 0:
				self.onRemove(None)
		else:
			evt.Skip()
		self._reset_timer()

	def onOpen(self, evt):
		self._reset_timer()
		if self.tabs.GetSelection() == 0:
			lst = self.listSaved
			idx = lst.GetFirstSelected()
			if idx == -1:
				return
			name = lst.GetItemText(idx, 0)
			path = self.manager._files.get(name)
		else:
			lst = self.listRecent
			idx = lst.GetFirstSelected()
			if idx == -1:
				return
			if idx >= len(self._displayedRecentPaths):
				return
			path = self._displayedRecentPaths[idx]

		if not path or not os.path.isfile(path):
			return

		self.manager.addToRecent(path)
		self._stop_timer()
		self._pendingOpenPath = path
		gui.mainFrame.postPopup()
		self.Close()

	def onAdd(self, evt):
		self._reset_timer()
		if not self.manager._newFile:
			ui.message(_("No file selected in Explorer to add."))
			return
		default = os.path.basename(self.manager._newFile)
		dlg = wx.TextEntryDialog(self, _("Enter display name"), TITLE, default)
		if dlg.ShowModal() == wx.ID_OK:
			name = dlg.GetValue().strip()
			if name:
				if name in self.manager._files:
					gui.messageBox(_("This name already exists."), TITLE, wx.OK | wx.ICON_WARNING)
				else:
					self.manager._files[name] = self.manager._newFile
					if name not in self.manager._order:
						self.manager._order.append(name)
					self.manager.saveConfig()
					self.updateFiles()
		dlg.Destroy()

	def onEdit(self, evt):
		self._reset_timer()
		if self.tabs.GetSelection() != 0:
			return
		idx = self.listSaved.GetFirstSelected()
		if idx == -1:
			return
		oldName = self.listSaved.GetItemText(idx, 0)
		dlg = wx.TextEntryDialog(self, _("Rename"), TITLE, oldName)
		if dlg.ShowModal() == wx.ID_OK:
			newName = dlg.GetValue().strip()
			if newName and newName != oldName:
				if newName in self.manager._files:
					gui.messageBox(_("This name already exists."), TITLE, wx.OK | wx.ICON_WARNING)
				else:
					path = self.manager._files.pop(oldName)
					self.manager._files[newName] = path
					self.manager._order = [newName if x == oldName else x for x in self.manager._order]
					if oldName in self.manager._pinned:
						self.manager._pinned.remove(oldName)
						self.manager._pinned.add(newName)
					self.manager.saveConfig()
					self.updateFiles()
		dlg.Destroy()

	def onRemove(self, evt):
		self._reset_timer()
		if self.tabs.GetSelection() != 0:
			return
		idx = self.listSaved.GetFirstSelected()
		if idx == -1:
			return
		name = self.listSaved.GetItemText(idx, 0)
		if gui.messageBox(_("Remove {}?").format(name), TITLE, wx.YES_NO) == wx.YES:
			self.manager._files.pop(name, None)
			if name in self.manager._order:
				self.manager._order.remove(name)
			self.manager._pinned.discard(name)
			self.manager.saveConfig()
			self.updateFiles()

	def onClearRecent(self, evt):
		self._reset_timer()
		if gui.messageBox(_("Clear history?"), TITLE, wx.YES_NO) == wx.YES:
			self.manager._recentFiles = []
			self.manager.saveConfig()
			self.updateFiles()

	def onContextMenu(self, evt):
		self._stop_timer()
		self._contextMenuOpen = True

		if self.tabs.GetSelection() == 0:
			lst = self.listSaved
			idx = lst.GetFirstSelected()
			if idx == -1:
				self._contextMenuOpen = False
				self._reset_timer()
				return
			name = lst.GetItemText(idx, 0)
			path = self.manager._files.get(name)
		else:
			lst = self.listRecent
			idx = lst.GetFirstSelected()
			if idx == -1:
				self._contextMenuOpen = False
				self._reset_timer()
				return
			if idx >= len(self._displayedRecentPaths):
				self._contextMenuOpen = False
				self._reset_timer()
				return
			path = self._displayedRecentPaths[idx]

		menu = wx.Menu()

		if path and os.path.isfile(path):
			ext = os.path.splitext(path)[1].lower()
			if ext in ('.exe', '.bat', '.cmd', '.msi'):
				itemAdmin = menu.Append(wx.ID_ANY, _("Run as Administrator"))
				self.Bind(wx.EVT_MENU, lambda e: self.runAsAdmin(path), itemAdmin)
				menu.AppendSeparator()

		if self.tabs.GetSelection() == 0:
			pin_label = _("Unpin") if name in self.manager._pinned else _("Pin to top")
			itemPin = menu.Append(wx.ID_ANY, pin_label)
			self.Bind(wx.EVT_MENU, lambda e: self.onTogglePin(name), itemPin)
			menu.AppendSeparator()
			itemEdit = menu.Append(wx.ID_ANY, _("Edit"))
			itemDelete = menu.Append(wx.ID_ANY, _("Delete"))

			if self.manager._sortMode == "CUSTOM":
				menu.AppendSeparator()
				itemUp = menu.Append(wx.ID_ANY, _("Move Up"))
				itemDown = menu.Append(wx.ID_ANY, _("Move Down"))
				self.Bind(wx.EVT_MENU, lambda e: self.moveItem(name, -1), itemUp)
				self.Bind(wx.EVT_MENU, lambda e: self.moveItem(name, 1), itemDown)

			self.Bind(wx.EVT_MENU, self.onEdit, itemEdit)
			self.Bind(wx.EVT_MENU, self.onRemove, itemDelete)
		else:
			itemDelete = menu.Append(wx.ID_ANY, _("Remove from Recent"))
			self.Bind(wx.EVT_MENU, lambda e, p=path: self.onRemoveRecentByPath(p), itemDelete)

		def on_menu_close(event):
			self._contextMenuOpen = False
			self._reset_timer()
			event.Skip()

		menu.Bind(wx.EVT_MENU_CLOSE, on_menu_close)
		lst.PopupMenu(menu)
		menu.Destroy()

	def onRemoveRecentByPath(self, path):
		self._reset_timer()
		if path in self.manager._recentFiles:
			if gui.messageBox(_("Remove {} from recent list?").format(os.path.basename(path)), TITLE, wx.YES_NO) == wx.YES:
				self.manager._recentFiles.remove(path)
				self.manager.saveConfig()
				self.updateFiles()

	def onTogglePin(self, name):
		self._reset_timer()
		if name in self.manager._pinned:
			self.manager._pinned.remove(name)
		else:
			self.manager._pinned.add(name)
		self.manager.saveConfig()
		self.updateFiles()

	def moveItem(self, targetName, direction):
		self._reset_timer()
		if self.tabs.GetSelection() != 0:
			return

		pinnedList = [x for x in self.manager._order if x in self.manager._pinned]
		unpinnedList = [x for x in self.manager._order if x not in self.manager._pinned and x in self.manager._files]

		if targetName in pinnedList:
			currentIndex = pinnedList.index(targetName)
			newIndex = currentIndex + direction
			if 0 <= newIndex < len(pinnedList):
				pinnedList[currentIndex], pinnedList[newIndex] = pinnedList[newIndex], pinnedList[currentIndex]
				self.manager._order = pinnedList + unpinnedList
				self.manager.saveConfig()
				self.updateFiles(newIndex)
		else:
			currentIndex = unpinnedList.index(targetName)
			newIndex = currentIndex + direction
			if 0 <= newIndex < len(unpinnedList):
				unpinnedList[currentIndex], unpinnedList[newIndex] = unpinnedList[newIndex], unpinnedList[currentIndex]
				self.manager._order = pinnedList + unpinnedList
				self.manager.saveConfig()
				self.updateFiles(len(pinnedList) + newIndex)

	def runAsAdmin(self, path):
		self._reset_timer()
		try:
			shell32 = ctypes.windll.shell32
			shell32.ShellExecuteW.argtypes = [
				wintypes.HWND,
				wintypes.LPCWSTR,
				wintypes.LPCWSTR,
				wintypes.LPCWSTR,
				wintypes.LPCWSTR,
				ctypes.c_int
			]
			shell32.ShellExecuteW.restype = wintypes.HINSTANCE
			shell32.ShellExecuteW(None, "runas", path, None, None, 1)
			self.manager.addToRecent(path)
			self.Close()
		except OSError as e:
			logHandler.log.warning(f"Failed to run as admin: {e}", exc_info=True)

	def updateFiles(self, selectIdx=0):
		self.listSaved.DeleteAllItems()
		self.listRecent.DeleteAllItems()
		self._displayedRecentPaths = []
		f_type = self.filterCombo.GetValue().lower()
		searchText = (self.searchSaved.GetValue() if self.tabs.GetSelection() == 0 else self.searchRecent.GetValue()).strip().lower()
		exts = {
			"audio": ('.mp3', '.wav', '.flac', '.m4a', '.ogg'),
			"video": ('.mp4', '.mkv', '.avi', '.mov'),
			"document": ('.pdf', '.docx', '.txt', '.xlsx', '.pptx'),
			"code": ('.py', '.cpp', '.java', '.js', '.html', '.css'),
			"exe": ('.exe', '.bat', '.cmd', '.msi')
		}
		if self.tabs.GetSelection() == 0:
			pinned = [x for x in self.manager._order if x in self.manager._pinned]
			unpinned = [x for x in self.manager._order if x not in self.manager._pinned and x in self.manager._files]
			if self.manager._sortMode == "UPPERCASE":
				unpinned.sort(key=lambda x: x.upper())
			elif self.manager._sortMode == "LOWERCASE":
				unpinned.sort(key=lambda x: x.lower(), reverse=True)
			items = pinned + unpinned
			count = 0
			for name in items:
				if searchText and searchText not in name.lower():
					continue
				path = self.manager._files[name]
				if f_type == "all" or path.lower().endswith(exts.get(f_type, ())):
					idx = self.listSaved.InsertItem(count, name)
					if self.manager._showPath:
						self.listSaved.SetItem(idx, 1, path)
					count += 1
			if self.listSaved.GetItemCount() > 0 and selectIdx < self.listSaved.GetItemCount():
				self.listSaved.Select(selectIdx)
				self.listSaved.Focus(selectIdx)
			has_selection = self.listSaved.GetFirstSelected() != -1
			self.btnEdit.Enable(has_selection)
			self.btnRemove.Enable(has_selection)
			self.btnAdd.Enable(bool(self.manager._newFile))
		else:
			count = 0
			for p in self.manager._recentFiles:
				name = os.path.basename(p)
				if searchText and searchText not in name.lower():
					continue
				if f_type == "all" or p.lower().endswith(exts.get(f_type, ())):
					idx = self.listRecent.InsertItem(count, name)
					self.listRecent.SetItem(idx, 1, p)
					self._displayedRecentPaths.append(p)
					count += 1



