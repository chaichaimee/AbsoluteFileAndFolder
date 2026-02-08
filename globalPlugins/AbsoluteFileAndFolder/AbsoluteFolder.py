# AbsoluteFolder.py

import os
import wx
import ui
import api
import gui
import globalVars
import json
import time
import ctypes
import urllib.parse
from comtypes.client import CreateObject as COMCreate
import addonHandler
import core

addonHandler.initTranslation()

_FF_JSON_FILE = os.path.abspath(os.path.join(globalVars.appArgs.configPath, "AbsoluteFolders.json"))
TITLE = _("Absolute Folders")

# Get system uptime in milliseconds
def get_system_uptime():
    try:
        lib = ctypes.windll.kernel32
        tick_count = lib.GetTickCount64()
        return tick_count() if hasattr(tick_count, '__call__') else tick_count
    except Exception:
        return 0

class AbsoluteFolderManager:
    def __init__(self):
        self._files = {}
        self._order = []
        self._pinned = set()
        self._showPath = False
        self._sortMode = "UPPERCASE"
        self._newFolder = ""
        self._autoLoadLastFolder = False
        self._lastOpenedFolders = []
        self._recentFolders = []
        self._lastSystemUptime = 0
        self._systemRestartDetected = False
        self.dialog = None

    def _getCurrentPathFromExplorer(self):
        """Get current opened folder path from Windows Explorer address bar using multiple reliable methods"""
        try:
            fg = api.getForegroundObject()
            if not fg or not fg.appModule or fg.appModule.appName != "explorer":
                return None

            shell = COMCreate("Shell.Application")
            if not shell:
                return None

            for window in shell.Windows():
                try:
                    if not window or window.hwnd != fg.windowHandle:
                        continue

                    # Method 1: Primary - Document.Folder.Self.Path
                    if hasattr(window, "Document") and window.Document:
                        folder = window.Document.Folder
                        if folder and hasattr(folder, "Self"):
                            path = folder.Self.Path
                            if path and os.path.isdir(path):
                                return os.path.normpath(path)

                    # Method 2: Fallback - LocationURL (file:/// protocol)
                    if hasattr(window, "LocationURL") and window.LocationURL:
                        url = window.LocationURL
                        if url.startswith("file:///"):
                            # Convert file:///C:/Users/... to C:\Users\...
                            path = urllib.parse.unquote(url[8:])  # remove file:///
                            path = path.replace("/", "\\")
                            if os.path.isdir(path):
                                return os.path.normpath(path)

                    # Method 3: Very last resort - LocationName if it's a full path
                    if hasattr(window, "LocationName") and window.LocationName:
                        possible_path = window.LocationName
                        if os.path.isabs(possible_path) and os.path.isdir(possible_path):
                            return os.path.normpath(possible_path)

                except Exception:
                    continue

            # Fallback using focus object instead of foreground
            focus = api.getFocusObject()
            if focus and focus.appModule and focus.appModule.appName == "explorer":
                for window in shell.Windows():
                    try:
                        if not window or window.hwnd != focus.windowHandle:
                            continue
                        if hasattr(window, "Document") and window.Document:
                            folder = window.Document.Folder
                            if folder and hasattr(folder, "Self"):
                                path = folder.Self.Path
                                if path and os.path.isdir(path):
                                    return os.path.normpath(path)
                    except Exception:
                        continue

        except Exception:
            pass

        return None

    def loadConfig(self):
        if os.path.isfile(_FF_JSON_FILE):
            try:
                with open(_FF_JSON_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self._sortMode = data.get("sortMode", "UPPERCASE")
                self._pinned = set(data.get("pinned", []))
                self._showPath = data.get("showPath", False)
                self._files = data.get("files", {})
                self._order = data.get("order", list(self._files.keys()))
                self._autoLoadLastFolder = data.get("autoLoadLastFolder", False)
                self._lastOpenedFolders = data.get("lastOpenedFolders", [])
                self._recentFolders = data.get("recentFolders", [])
                self._lastSystemUptime = data.get("lastSystemUptime", 0)
            except Exception:
                pass

    def saveConfig(self):
        data = {
            "sortMode": self._sortMode,
            "files": self._files,
            "order": self._order,
            "pinned": list(self._pinned),
            "showPath": self._showPath,
            "autoLoadLastFolder": self._autoLoadLastFolder,
            "lastOpenedFolders": self._lastOpenedFolders,
            "recentFolders": self._recentFolders,
            "lastSystemUptime": self._lastSystemUptime
        }
        try:
            with open(_FF_JSON_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def shouldAutoOpenOnStartup(self):
        """Check if folders should be opened based on Windows restart only."""
        if not self._autoLoadLastFolder or not self._lastOpenedFolders:
            return False
        
        current_uptime = get_system_uptime()
        
        # If we have no previous uptime recorded (first run), don't open folders
        if self._lastSystemUptime == 0:
            self._lastSystemUptime = current_uptime
            self.saveConfig()
            return False
        
        # Detect Windows restart: current uptime is less than previous uptime
        # This means the system was rebooted
        if current_uptime < self._lastSystemUptime:
            self._lastSystemUptime = current_uptime
            self.saveConfig()
            return True
        
        # If current uptime is very small (less than 30 seconds) and we haven't recorded it yet
        # This could be a fresh Windows boot
        if current_uptime < 30000 and self._lastSystemUptime > 30000:
            self._lastSystemUptime = current_uptime
            self.saveConfig()
            return True
        
        # Update uptime but don't open folders for NVDA restarts
        self._lastSystemUptime = current_uptime
        self.saveConfig()
        return False

    def addToRecent(self, path):
        if path and os.path.isdir(path):
            if path in self._recentFolders:
                self._recentFolders.remove(path)
            self._recentFolders.insert(0, path)
            self._recentFolders = self._recentFolders[:20]
            self.saveConfig()

    def show(self):
        """Show the dialog. Can be called from anywhere. If in Explorer, Add button will be enabled."""
        self.loadConfig()
        path = self._getCurrentPathFromExplorer()

        if path and os.path.isdir(path):
            self._newFolder = path
        else:
            self._newFolder = ""

        self.dialog = AbsoluteFoldersDialog(gui.mainFrame, self)
        gui.mainFrame.prePopup()
        self.dialog.Show()
        self.dialog.CentreOnScreen()
        gui.mainFrame.postPopup()

class AbsoluteFoldersDialog(wx.Dialog):
    def __init__(self, parent, manager):
        super().__init__(parent, title=TITLE, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER | wx.MAXIMIZE_BOX)
        self.manager = manager
        self._initUI()
        self._bindEvents()
        self.updateFiles()
        self.updateAutoOpenList()
        wx.CallAfter(self.listSaved.SetFocus)

    def _initUI(self):
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        self.tabs = wx.Notebook(self)
        self.panelSaved = wx.Panel(self.tabs)
        self.panelRecent = wx.Panel(self.tabs)
        self.tabs.AddPage(self.panelSaved, _("Saved Folders"))
        self.tabs.AddPage(self.panelRecent, _("Recent Folders"))
        mainSizer.Add(self.tabs, 1, wx.EXPAND | wx.ALL, 5)

        savedSizer = wx.BoxSizer(wx.VERTICAL)
        self.listSaved = wx.ListCtrl(self.panelSaved, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.BORDER_SUNKEN)
        self.listSaved.InsertColumn(0, _("Name"), width=250)
        self.listSaved.InsertColumn(1, _("Path"), width=400)
        savedSizer.Add(self.listSaved, 1, wx.EXPAND | wx.ALL, 5)
        self.panelSaved.SetSizer(savedSizer)

        recentSizer = wx.BoxSizer(wx.VERTICAL)
        self.listRecent = wx.ListCtrl(self.panelRecent, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.BORDER_SUNKEN)
        self.listRecent.InsertColumn(0, _("Folder Name"), width=250)
        self.listRecent.InsertColumn(1, _("Path"), width=400)
        recentSizer.Add(self.listRecent, 1, wx.EXPAND | wx.ALL, 5)
        self.btnClearRecent = wx.Button(self.panelRecent, label=_("Clear History"))
        recentSizer.Add(self.btnClearRecent, 0, wx.ALIGN_RIGHT | wx.ALL, 5)
        self.panelRecent.SetSizer(recentSizer)

        # Button sizer for Saved Folders tab - placed at top
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.btnAdd = wx.Button(self, label=_("&Add"))
        self.btnOpen = wx.Button(self, label=_("&Open"))
        self.btnEdit = wx.Button(self, label=_("&Edit"))
        self.btnRemove = wx.Button(self, label=_("&Remove"))
        btnSizer.Add(self.btnAdd, 0, wx.RIGHT, 5)
        btnSizer.Add(self.btnOpen, 0, wx.RIGHT, 5)
        btnSizer.Add(self.btnEdit, 0, wx.RIGHT, 5)
        btnSizer.Add(self.btnRemove, 0, wx.RIGHT, 5)
        mainSizer.Add(btnSizer, 0, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, 10)

        optionsSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.chkShowPath = wx.CheckBox(self, label=_("&Show paths"))
        self.chkShowPath.SetValue(self.manager._showPath)
        optionsSizer.Add(self.chkShowPath, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        choices = [_("Custom order"), _("Ascending, a-z"), _("Descending z-a")]
        self.sortCombo = wx.ComboBox(self, choices=choices, style=wx.CB_READONLY)
        mode_map = {"CUSTOM": 0, "UPPERCASE": 1, "LOWERCASE": 2}
        self.sortCombo.SetSelection(mode_map.get(self.manager._sortMode, 1))
        optionsSizer.Add(self.sortCombo, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        mainSizer.Add(optionsSizer, 0, wx.EXPAND | wx.ALL, 5)

        # Auto-open section
        autoOpenSizer = wx.BoxSizer(wx.VERTICAL)
        self.chkAutoLoad = wx.CheckBox(self, label=_("Remember and open folders automatically on restart"))
        self.chkAutoLoad.SetValue(self.manager._autoLoadLastFolder)
        self.chkAutoLoad.Bind(wx.EVT_CHECKBOX, self.onAutoLoadChanged)
        autoOpenSizer.Add(self.chkAutoLoad, 0, wx.ALL, 5)

        # Auto-open folders list (hidden by default)
        self.autoOpenPanel = wx.Panel(self)
        autoOpenListSizer = wx.BoxSizer(wx.VERTICAL)
        
        self.listAutoOpen = wx.ListCtrl(self.autoOpenPanel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.BORDER_SUNKEN)
        self.listAutoOpen.InsertColumn(0, _("Folder Name"), width=250)
        self.listAutoOpen.InsertColumn(1, _("Path"), width=400)
        autoOpenListSizer.Add(self.listAutoOpen, 1, wx.EXPAND | wx.ALL, 5)
        
        self.autoOpenPanel.SetSizer(autoOpenListSizer)
        autoOpenSizer.Add(self.autoOpenPanel, 0, wx.EXPAND | wx.ALL, 5)
        
        # Show/hide the auto-open panel based on initial state
        if not self.manager._autoLoadLastFolder:
            self.autoOpenPanel.Hide()
        
        mainSizer.Add(autoOpenSizer, 0, wx.EXPAND | wx.ALL, 5)

        # Close button at the bottom
        closeBtnSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.btnClose = wx.Button(self, wx.ID_CLOSE)
        closeBtnSizer.AddStretchSpacer()
        closeBtnSizer.Add(self.btnClose, 0)
        mainSizer.Add(closeBtnSizer, 0, wx.EXPAND | wx.ALL, 10)

        self.SetSizer(mainSizer)
        self.SetMinSize((800, 600))
        self.Fit()

    def _bindEvents(self):
        self.tabs.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.onTabChanged)
        self.chkShowPath.Bind(wx.EVT_CHECKBOX, self.onShowPathChanged)
        self.sortCombo.Bind(wx.EVT_COMBOBOX, self.onSortChanged)
        self.btnAdd.Bind(wx.EVT_BUTTON, self.onAdd)
        self.btnOpen.Bind(wx.EVT_BUTTON, self.onOpen)
        self.btnEdit.Bind(wx.EVT_BUTTON, self.onEdit)
        self.btnRemove.Bind(wx.EVT_BUTTON, self.onRemove)
        self.btnClose.Bind(wx.EVT_BUTTON, lambda e: self.Close())
        self.btnClearRecent.Bind(wx.EVT_BUTTON, self.onClearRecent)
        self.listSaved.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.onOpen)
        self.listRecent.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.onOpen)
        self.listSaved.Bind(wx.EVT_CONTEXT_MENU, self.onContextMenu)
        self.listSaved.Bind(wx.EVT_KEY_DOWN, self.onKeyDown)
        self.listRecent.Bind(wx.EVT_KEY_DOWN, self.onKeyDown)
        self.listAutoOpen.Bind(wx.EVT_CONTEXT_MENU, self.onAutoOpenContextMenu)
        self.listAutoOpen.Bind(wx.EVT_KEY_DOWN, self.onAutoOpenKeyDown)
        self.Bind(wx.EVT_CHAR_HOOK, self.onCharHook)

    def onAutoLoadChanged(self, evt):
        self.manager._autoLoadLastFolder = self.chkAutoLoad.GetValue()
        self.manager.saveConfig()
        if self.manager._autoLoadLastFolder:
            self.autoOpenPanel.Show()
        else:
            self.autoOpenPanel.Hide()
        self.Layout()
        self.Fit()

    def onAutoOpenContextMenu(self, evt):
        idx = self.listAutoOpen.GetFirstSelected()
        if idx == -1:
            return
        
        menu = wx.Menu()
        itemDelete = menu.Append(wx.ID_ANY, _("Delete"))
        self.Bind(wx.EVT_MENU, self.onAutoOpenDelete, itemDelete)
        self.listAutoOpen.PopupMenu(menu)
        menu.Destroy()

    def onAutoOpenKeyDown(self, evt):
        if evt.GetKeyCode() == wx.WXK_DELETE:
            self.onAutoOpenDelete(None)
        else:
            evt.Skip()

    def onAutoOpenDelete(self, evt):
        idx = self.listAutoOpen.GetFirstSelected()
        if idx == -1:
            return
        
        if idx < len(self.manager._lastOpenedFolders):
            folder_path = self.manager._lastOpenedFolders[idx]
            if gui.messageBox(_("Remove {} from auto-open list?").format(folder_path), TITLE, wx.YES_NO) == wx.YES:
                self.manager._lastOpenedFolders.pop(idx)
                self.manager.saveConfig()
                self.updateAutoOpenList()

    def onTabChanged(self, evt):
        self.updateFiles()
        # Enable/disable buttons based on selected tab
        is_saved_tab = self.tabs.GetSelection() == 0
        self.btnAdd.Enable(is_saved_tab and bool(self.manager._newFolder))
        self.btnEdit.Enable(is_saved_tab)
        self.btnRemove.Enable(is_saved_tab)
        if is_saved_tab:
            self.listSaved.SetFocus()
        else:
            self.listRecent.SetFocus()

    def onCharHook(self, evt):
        if evt.GetKeyCode() == wx.WXK_ESCAPE:
            self.Close()
        else:
            evt.Skip()

    def onKeyDown(self, evt):
        if evt.GetKeyCode() == wx.WXK_DELETE:
            if self.tabs.GetSelection() == 0:  # Saved Folders tab
                self.onRemove(None)
        else:
            evt.Skip()

    def onShowPathChanged(self, evt):
        self.manager._showPath = self.chkShowPath.GetValue()
        self.manager.saveConfig()
        self.updateFiles()

    def onSortChanged(self, evt):
        idx = self.sortCombo.GetSelection()
        self.manager._sortMode = ["CUSTOM", "UPPERCASE", "LOWERCASE"][idx]
        self.manager.saveConfig()
        self.updateFiles()

    def onAdd(self, evt):
        if not self.manager._newFolder:
            ui.message(_("No folder selected in Explorer to add."))
            return
        default = os.path.basename(self.manager._newFolder)
        dlg = wx.TextEntryDialog(self, _("Enter display name"), TITLE, default)
        if dlg.ShowModal() == wx.ID_OK:
            name = dlg.GetValue().strip()
            if name:
                if name in self.manager._files:
                    gui.messageBox(_("This name already exists."), TITLE, wx.OK | wx.ICON_WARNING)
                else:
                    self.manager._files[name] = self.manager._newFolder
                    if name not in self.manager._order:
                        self.manager._order.append(name)
                    # Add to auto-open list if checkbox is checked
                    if self.manager._autoLoadLastFolder and self.manager._newFolder not in self.manager._lastOpenedFolders:
                        self.manager._lastOpenedFolders.append(self.manager._newFolder)
                    self.manager.saveConfig()
                    self.updateFiles()
                    self.updateAutoOpenList()
        dlg.Destroy()

    def onOpen(self, evt):
        if self.tabs.GetSelection() == 0:  # Saved Folders tab
            lst = self.listSaved
            idx = lst.GetFirstSelected()
            if idx == -1:
                return
            name = lst.GetItemText(idx, 0)
            path = self.manager._files.get(name)
        else:  # Recent Folders tab
            lst = self.listRecent
            idx = lst.GetFirstSelected()
            if idx == -1:
                return
            path = self.manager._recentFolders[idx]
        
        if path and os.path.isdir(path):
            os.startfile(path)
            self.manager.addToRecent(path)
            # Add to auto-open list if checkbox is checked
            if self.manager._autoLoadLastFolder and path not in self.manager._lastOpenedFolders:
                self.manager._lastOpenedFolders.append(path)
                self.manager.saveConfig()
                self.updateAutoOpenList()
            self.Close()

    def onEdit(self, evt):
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
        if self.tabs.GetSelection() != 0:
            return
        idx = self.listSaved.GetFirstSelected()
        if idx == -1:
            return
        name = self.listSaved.GetItemText(idx, 0)
        if gui.messageBox(_("Remove {}?").format(name), TITLE, wx.YES_NO) == wx.YES:
            path = self.manager._files.get(name)
            self.manager._files.pop(name, None)
            if name in self.manager._order:
                self.manager._order.remove(name)
            self.manager._pinned.discard(name)
            # Remove from auto-open list if present
            if path and path in self.manager._lastOpenedFolders:
                self.manager._lastOpenedFolders.remove(path)
            self.manager.saveConfig()
            self.updateFiles()
            self.updateAutoOpenList()

    def onContextMenu(self, evt):
        if self.tabs.GetSelection() != 0:
            return
        idx = self.listSaved.GetFirstSelected()
        if idx == -1:
            return
        name = self.listSaved.GetItemText(idx, 0)
        menu = wx.Menu()
        pin_label = _("Unpin") if name in self.manager._pinned else _("Pin to top")
        itemPin = menu.Append(wx.ID_ANY, pin_label)
        menu.AppendSeparator()
        itemEdit = menu.Append(wx.ID_ANY, _("Edit"))
        itemDelete = menu.Append(wx.ID_ANY, _("Delete"))
        if self.manager._sortMode == "CUSTOM":
            menu.AppendSeparator()
            itemUp = menu.Append(wx.ID_ANY, _("Move Up"))
            itemDown = menu.Append(wx.ID_ANY, _("Move Down"))
            self.Bind(wx.EVT_MENU, lambda e: self.moveItem(-1), itemUp)
            self.Bind(wx.EVT_MENU, lambda e: self.moveItem(1), itemDown)
        self.Bind(wx.EVT_MENU, lambda e: self.onTogglePin(name), itemPin)
        self.Bind(wx.EVT_MENU, self.onEdit, itemEdit)
        self.Bind(wx.EVT_MENU, self.onRemove, itemDelete)
        self.listSaved.PopupMenu(menu)
        menu.Destroy()

    def onTogglePin(self, name):
        if name in self.manager._pinned:
            self.manager._pinned.remove(name)
        else:
            self.manager._pinned.add(name)
        self.manager.saveConfig()
        self.updateFiles()

    def moveItem(self, direction):
        idx = self.listSaved.GetFirstSelected()
        if idx == -1:
            return
        name = self.listSaved.GetItemText(idx, 0)
        if name in self.manager._pinned:
            return
        unpinned = [x for x in self.manager._order if x not in self.manager._pinned]
        cur_idx = unpinned.index(name)
        new_idx = cur_idx + direction
        if 0 <= new_idx < len(unpinned):
            unpinned[cur_idx], unpinned[new_idx] = unpinned[new_idx], unpinned[cur_idx]
            self.manager._order = sorted([x for x in self.manager._order if x in self.manager._pinned]) + unpinned
            self.manager.saveConfig()
            self.updateFiles(new_idx + len(self.manager._pinned))

    def onClearRecent(self, evt):
        if gui.messageBox(_("Clear history?"), TITLE, wx.YES_NO) == wx.YES:
            self.manager._recentFolders = []
            self.manager.saveConfig()
            self.updateFiles()

    def updateFiles(self, selectIdx=0):
        self.listSaved.DeleteAllItems()
        self.listRecent.DeleteAllItems()
        if self.tabs.GetSelection() == 0:
            pinned = sorted([x for x in self.manager._order if x in self.manager._pinned], key=lambda x: x.upper())
            unpinned = [x for x in self.manager._order if x not in self.manager._pinned and x in self.manager._files]
            if self.manager._sortMode == "UPPERCASE":
                unpinned.sort(key=lambda x: x.upper())
            elif self.manager._sortMode == "LOWERCASE":
                unpinned.sort(key=lambda x: x.lower(), reverse=True)
            items = pinned + unpinned
            for i, name in enumerate(items):
                self.listSaved.InsertItem(i, name)
                if self.manager._showPath:
                    self.listSaved.SetItem(i, 1, self.manager._files[name])
            if self.listSaved.GetItemCount() > 0:
                self.listSaved.Select(selectIdx)
            # Update button states
            has_selection = self.listSaved.GetFirstSelected() != -1
            self.btnEdit.Enable(has_selection)
            self.btnRemove.Enable(has_selection)
            self.btnAdd.Enable(bool(self.manager._newFolder))
        else:
            for i, path in enumerate(self.manager._recentFolders):
                self.listRecent.InsertItem(i, os.path.basename(path))
                if self.manager._showPath:
                    self.listRecent.SetItem(i, 1, path)

    def updateAutoOpenList(self):
        self.listAutoOpen.DeleteAllItems()
        for i, path in enumerate(self.manager._lastOpenedFolders):
            if os.path.isdir(path):
                self.listAutoOpen.InsertItem(i, os.path.basename(path))
                self.listAutoOpen.SetItem(i, 1, path)