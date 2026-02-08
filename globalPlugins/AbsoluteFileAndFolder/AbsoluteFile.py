# AbsoluteFile.py

import os
import wx
import ui
import api
import gui
import globalVars
import json
import addonHandler
import core
import ctypes
from comtypes.client import CreateObject as COMCreate
import urllib.parse

addonHandler.initTranslation()

_AF_JSON_FILE = os.path.abspath(os.path.join(globalVars.appArgs.configPath, "AbsoluteFiles.json"))
TITLE = _("Absolute Files")

class AbsoluteFileManager:
    def __init__(self):
        self._files = {}
        self._order = []
        self._pinned = set()
        self._recentFiles = []
        self._showPath = False
        self._sortMode = "UPPERCASE"
        self._newFile = ""
        self.loadConfig()

    def _getCurrentPathFromExplorer(self):
        """Get current file path from Windows Explorer if a file is selected"""
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

                    # Method 1: Try to get focused item
                    if hasattr(window, "Document") and window.Document:
                        try:
                            item = window.Document.FocusedItem
                            if item:
                                path = item.Path
                                if path and os.path.isfile(path):
                                    return os.path.normpath(path)
                        except Exception:
                            pass

                    # Method 2: Try LocationURL
                    if hasattr(window, "LocationURL") and window.LocationURL:
                        url = window.LocationURL
                        if url.startswith("file:///"):
                            # Convert file:///C:/Users/... to C:\Users\...
                            path = urllib.parse.unquote(url[8:])  # remove file:///
                            path = path.replace("/", "\\")
                            if os.path.isfile(path):
                                return os.path.normpath(path)

                except Exception:
                    continue

            # Fallback using focus object
            focus = api.getFocusObject()
            if focus and focus.appModule and focus.appModule.appName == "explorer":
                for window in shell.Windows():
                    try:
                        if not window or window.hwnd != focus.windowHandle:
                            continue
                        if hasattr(window, "Document") and window.Document:
                            try:
                                item = window.Document.FocusedItem
                                if item:
                                    path = item.Path
                                    if path and os.path.isfile(path):
                                        return os.path.normpath(path)
                            except Exception:
                                pass
                    except Exception:
                        continue

        except Exception:
            pass

        return None

    def loadConfig(self):
        if os.path.isfile(_AF_JSON_FILE):
            try:
                with open(_AF_JSON_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self._files = data.get("files", {})
                self._order = data.get("order", list(self._files.keys()))
                self._pinned = set(data.get("pinned", []))
                self._recentFiles = data.get("recentFiles", [])
                self._showPath = data.get("showPath", False)
                self._sortMode = data.get("sortMode", "UPPERCASE")
            except Exception:
                pass

    def saveConfig(self):
        data = {
            "files": self._files,
            "order": self._order,
            "pinned": list(self._pinned),
            "recentFiles": self._recentFiles,
            "showPath": self._showPath,
            "sortMode": self._sortMode
        }
        try:
            with open(_AF_JSON_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def addToRecent(self, path):
        if path and os.path.isfile(path):
            if path in self._recentFiles:
                self._recentFiles.remove(path)
            self._recentFiles.insert(0, path)
            self._recentFiles = self._recentFiles[:20]
            self.saveConfig()

    def show(self):
        """Show the dialog. Can be called from anywhere. If in Explorer with file selected, Add button will be enabled."""
        self.loadConfig()
        path = self._getCurrentPathFromExplorer()
        
        if path and os.path.isfile(path):
            self._newFile = path
        elif path and os.path.isdir(path):
            # Don't show message, just disable add button
            self._newFile = ""
        else:
            self._newFile = ""
            
        self.dialog = AbsoluteFilesDialog(gui.mainFrame, self)
        gui.mainFrame.prePopup()
        self.dialog.Show()
        self.dialog.CentreOnScreen()
        gui.mainFrame.postPopup()

class AbsoluteFilesDialog(wx.Dialog):
    def __init__(self, parent, manager):
        super().__init__(parent, title=TITLE, style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER | wx.MAXIMIZE_BOX)
        self.manager = manager
        self._initUI()
        self._bindEvents()
        self.updateFiles()
        wx.CallAfter(self.listSaved.SetFocus)

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

        # Saved Files panel
        savedSizer = wx.BoxSizer(wx.VERTICAL)
        self.listSaved = wx.ListCtrl(self.panelSaved, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.BORDER_SUNKEN)
        self.listSaved.InsertColumn(0, _("Name"), width=250)
        self.listSaved.InsertColumn(1, _("Path"), width=400)
        savedSizer.Add(self.listSaved, 1, wx.EXPAND | wx.ALL, 5)
        
        # Button sizer for Saved Files tab
        savedBtnSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.btnAdd = wx.Button(self.panelSaved, label=_("&Add"))
        self.btnEdit = wx.Button(self.panelSaved, label=_("&Edit"))
        self.btnRemove = wx.Button(self.panelSaved, label=_("&Remove"))
        savedBtnSizer.Add(self.btnAdd, 0, wx.RIGHT, 5)
        savedBtnSizer.Add(self.btnEdit, 0, wx.RIGHT, 5)
        savedBtnSizer.Add(self.btnRemove, 0)
        savedSizer.Add(savedBtnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 5)
        
        self.panelSaved.SetSizer(savedSizer)

        # Recent Files panel
        recentSizer = wx.BoxSizer(wx.VERTICAL)
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

        # Main dialog buttons
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
        self.filterCombo.Bind(wx.EVT_COMBOBOX, lambda e: self.updateFiles())
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

    def onTabChanged(self, evt):
        self.updateFiles()
        if self.tabs.GetSelection() == 0:
            self.listSaved.SetFocus()
        else:
            self.listRecent.SetFocus()

    def onSortChanged(self, evt):
        idx = self.sortCombo.GetSelection()
        self.manager._sortMode = ["CUSTOM", "UPPERCASE", "LOWERCASE"][idx]
        self.manager.saveConfig()
        self.updateFiles()

    def onCharHook(self, evt):
        if evt.GetKeyCode() == wx.WXK_ESCAPE:
            self.Close()
        else:
            evt.Skip()

    def onKeyDown(self, evt):
        if evt.GetKeyCode() == wx.WXK_DELETE:
            if self.tabs.GetSelection() == 0:
                self.onRemove(None)
        else:
            evt.Skip()

    def onOpen(self, evt):
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
            path = self.manager._recentFiles[idx]
            
        if path and os.path.isfile(path):
            os.startfile(path)
            self.manager.addToRecent(path)
            self.Close()

    def onAdd(self, evt):
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
            self.manager._files.pop(name, None)
            if name in self.manager._order:
                self.manager._order.remove(name)
            self.manager._pinned.discard(name)
            self.manager.saveConfig()
            self.updateFiles()

    def onClearRecent(self, evt):
        if gui.messageBox(_("Clear history?"), TITLE, wx.YES_NO) == wx.YES:
            self.manager._recentFiles = []
            self.manager.saveConfig()
            self.updateFiles()

    def onContextMenu(self, evt):
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
            path = self.manager._recentFiles[idx]
            
        menu = wx.Menu()
        
        # Add "Run as Administrator" for executable files
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
                self.Bind(wx.EVT_MENU, lambda e: self.moveItem(-1), itemUp)
                self.Bind(wx.EVT_MENU, lambda e: self.moveItem(1), itemDown)
                
            self.Bind(wx.EVT_MENU, self.onEdit, itemEdit)
            self.Bind(wx.EVT_MENU, self.onRemove, itemDelete)
        else:
            itemDelete = menu.Append(wx.ID_ANY, _("Remove from Recent"))
            self.Bind(wx.EVT_MENU, self.onRemoveRecent, itemDelete)
            
        lst.PopupMenu(menu)
        menu.Destroy()

    def onRemoveRecent(self, evt):
        if self.tabs.GetSelection() != 1:
            return
        idx = self.listRecent.GetFirstSelected()
        if idx == -1:
            return
        if idx < len(self.manager._recentFiles):
            path = self.manager._recentFiles[idx]
            if gui.messageBox(_("Remove {} from recent list?").format(os.path.basename(path)), TITLE, wx.YES_NO) == wx.YES:
                self.manager._recentFiles.pop(idx)
                self.manager.saveConfig()
                self.updateFiles()

    def onTogglePin(self, name):
        if name in self.manager._pinned:
            self.manager._pinned.remove(name)
        else:
            self.manager._pinned.add(name)
        self.manager.saveConfig()
        self.updateFiles()

    def moveItem(self, direction):
        if self.tabs.GetSelection() != 0:
            return
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

    def runAsAdmin(self, path):
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", path, None, None, 1)
            self.manager.addToRecent(path)
            self.Close()
        except Exception:
            pass

    def updateFiles(self, selectIdx=0):
        self.listSaved.DeleteAllItems()
        self.listRecent.DeleteAllItems()
        f_type = self.filterCombo.GetValue().lower()
        exts = {
            "audio": ('.mp3', '.wav', '.flac', '.m4a', '.ogg'),
            "video": ('.mp4', '.mkv', '.avi', '.mov'),
            "document": ('.pdf', '.docx', '.txt', '.xlsx', '.pptx'),
            "code": ('.py', '.cpp', '.java', '.js', '.html', '.css'),
            "exe": ('.exe', '.bat', '.cmd', '.msi')
        }
        
        if self.tabs.GetSelection() == 0:
            pinned = sorted([x for x in self.manager._order if x in self.manager._pinned], key=lambda x: x.upper())
            unpinned = [x for x in self.manager._order if x not in self.manager._pinned and x in self.manager._files]
            if self.manager._sortMode == "UPPERCASE":
                unpinned.sort(key=lambda x: x.upper())
            elif self.manager._sortMode == "LOWERCASE":
                unpinned.sort(key=lambda x: x.lower(), reverse=True)
            items = pinned + unpinned
            count = 0
            for name in items:
                path = self.manager._files[name]
                if f_type == "all" or path.lower().endswith(exts.get(f_type, ())):
                    idx = self.listSaved.InsertItem(count, name)
                    if self.manager._showPath:
                        self.listSaved.SetItem(idx, 1, path)
                    count += 1
            if self.listSaved.GetItemCount() > 0:
                self.listSaved.Select(selectIdx)
            
            # Update button states for Saved Files tab
            has_selection = self.listSaved.GetFirstSelected() != -1
            self.btnEdit.Enable(has_selection)
            self.btnRemove.Enable(has_selection)
            self.btnAdd.Enable(bool(self.manager._newFile))
        else:
            count = 0
            for p in self.manager._recentFiles:
                if f_type == "all" or p.lower().endswith(exts.get(f_type, ())):
                    idx = self.listRecent.InsertItem(count, os.path.basename(p))
                    self.listRecent.SetItem(idx, 1, p)
                    count += 1