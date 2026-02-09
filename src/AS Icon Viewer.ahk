; AS Icon Viewer
; 09/02/2026

; Mesut Akcan
; -----------
; mesutakcan.blogspot.com
; github.com/akcansoft
; youtube.com/mesutakcan

#Requires AutoHotkey v2
#SingleInstance Force

#Include Gdip.ahk ; GDI+
#Include SaveFileDialog.ahk ; Custom Save File Dialog

; Global variables
A_ScriptName := "AS Icon Viewer v1.2"
global SettingsFile := A_ScriptDir "\saved_files.txt" ; For storing user-added files in the left panel
global FavoritesFile := A_ScriptDir "\favorites.txt" ; For storing favorite icons
global CurrentDllPath := "" ; Currently loaded DLL/EXE/ICO file path
global CurrentViewMode := 3 ; 0:SmallReport, 1:LargeReport, 2:SmallIcon, 3:LargeIcon
global IL_Small := 0 ; Small ImageList
global IL_Large := 0 ; Large ImageList
global lv_Favorites := 0 ; Favorites ListView
global IL_Favorites := 0 ; Favorites ImageList
global FavoriteIcons := [] ; Array of { path, num }
global dllFileListChanged := false ; To track changes in the DLL file list for saving
global FavoritesListChanged := false ; To track changes in the favorites list for saving
global PreviewIcon := { Path: "", Num: 0 } ; Currently previewed icon info
global ExportIconSize := 128 ; Default export size
global LastGuiW := 0 ; For tracking GUI size changes
global LastGuiH := 0

global Txt := {
	Remove: "➖ Remove",
	Add: "➕ Add",
	Clear: "🧹 Clear",
	Save: "💾 Save",
	Test: "🧪 Test icon",
	Copy: "📋 Copy",
	CopyCode: "📋 Copy Code",
	AddFile: "➕ &Add File ",
	RemoveFile: "➖ &Remove File",
	ClearList: "🧹 &Clear List",
	Refresh: "&Refresh",
	SwitchView: "&Switch View",
	CopyImage: "📋 &Copy Image",
	SaveImage: "💾 &Save Image ",
	AddToFavorites: "➕ &Add to Favorites",
	RemoveFromFavorites: "➖ &Remove from Favorites"
}

; ========== TRAY ICON & MENU ==========
try TraySetIcon(A_ScriptDir "\app_icon.ico")
Tray := A_TrayMenu
Tray.Delete() ; Delete the standard items.
Tray.Add "Open " A_ScriptName, (*) => mGui.Show()
Tray.Add() ; separator
Tray.AddStandard()
Tray.Default := "Open " A_ScriptName

; Create main GUI
mGui := Gui("+Resize +MinSize900x450", A_ScriptName)
mGui.SetFont("s9", "Segoe UI")

; ========== MENU BAR ==========
; File Menu
mnu_File := Menu()
mnu_File.Add(Txt.AddFile, (*) => AddCustomFile())
mnu_File.Add(Txt.RemoveFile, (*) => RemoveCustomFile())
mnu_File.Add(Txt.ClearList, (*) => ClearFileList())
mnu_File.Add() ; Separator
mnu_File.Add("⏻ E&xit", (*) => CloseApplication())

; View Menu
mnu_View := Menu()
mnu_View.Add(Txt.SwitchView, (*) => SwitchView())
mnu_View.Add(Txt.Refresh, (*) => LoadIcons())

; Icon Menu
mnu_Icon := Menu()
mnu_Icon.Add(Txt.CopyImage, (*) => CopyIconToClipboard())
mnu_Icon.Add(Txt.SaveImage, (*) => SaveIconToFile())
mnu_Icon.Add(Txt.CopyCode, (*) => CopyCurrentCode())
mnu_Icon.Add(Txt.Test, (*) => TestIcon())

; Favorites Menu
mnu_Favorites := Menu()
mnu_Favorites.Add(Txt.AddToFavorites, (*) => AddToFavorites())
mnu_Favorites.Add(Txt.RemoveFromFavorites, (*) => RemoveFromFavorites())

; Help Menu
mnu_Help := Menu()
mnu_Help.Add("&About ", (*) => ShowAbout())
mnu_Help.Add() ; Separator
mnu_Help.Add("Visit &Website", (*) => Run("https://mesutakcan.blogspot.com"))
mnu_Help.Add("&GitHub Repository", (*) => Run("https://github.com/akcansoft/AS-Icon-Viewer"))

; Main Menu Bar
mnu_Main := MenuBar()
mnu_Main.Add("&File", mnu_File)
mnu_Main.Add("&View", mnu_View)
mnu_Main.Add("F&avorites", mnu_Favorites)
mnu_Main.Add("&Icon", mnu_Icon)
mnu_Main.Add("&Help", mnu_Help)
mGui.MenuBar := mnu_Main

; ========== LEFT PANEL (x10-260) ==========
mGui.AddText("x10 y10 w250", "📁 Icon Files:")
btn_AddFile := mGui.AddButton("x10 y35 w80 h30", Txt.Add) ; Add button
btn_AddFile.OnEvent("Click", AddCustomFile)
btn_RemoveFile := mGui.AddButton("x95 y35 w80 h30", Txt.Remove) ; Remove button
btn_RemoveFile.OnEvent("Click", RemoveCustomFile)
btn_ClearList := mGui.AddButton("x180 y35 w80 h30", Txt.Clear) ; Clear List button
btn_ClearList.OnEvent("Click", ClearFileList)
lst_Files := mGui.AddListBox("x10 y70 w250 h465 +VScroll +HScroll", []) ; List of DLL/EXE/ICO files

; ========== MIDDLE PANEL (x270-740) ==========
mGui.AddText("x270 y10 w300", "🎨 Icons:")

mGui.SetFont("s11", "Segoe MDL2 Assets")
btn_Switch := mGui.AddButton("x600 y5 w140 h30", Chr(0xE8EB) " Switch View") ; Switch View button
btn_Switch.OnEvent("Click", SwitchView)

mGui.SetFont("s9", "Segoe UI")

; Icons ListView
lv_Icons := mGui.AddListView("x270 y35 w400 h470 Grid -Multi", ["Icon and Number"])
lv_Icons.ModifyCol(1, 380)

; Right-click menu for Icons (Item selected)
mnu_IconContext := Menu()
mnu_IconContext.Add(Txt.CopyImage, (*) => CopyIconToClipboard())
mnu_IconContext.Add(Txt.SaveImage, (*) => SaveIconToFile())
mnu_IconContext.Add() ; Separator
mnu_IconContext.Add(Txt.CopyCode, (*) => CopyCurrentCode())
mnu_IconContext.Add(Txt.Test, (*) => TestIcon())
mnu_IconContext.Add() ; Separator
mnu_IconContext.Add(Txt.AddToFavorites, (*) => AddToFavorites())
mnu_IconContext.Add(Txt.Refresh, (*) => LoadIcons())
mnu_IconContext.Add(Txt.SwitchView, (*) => SwitchView())

; Right-click menu for Icons (Empty space)
mnu_IconEmptyContext := Menu()
mnu_IconEmptyContext.Add(Txt.Refresh, (*) => LoadIcons())
mnu_IconEmptyContext.Add(Txt.SwitchView, (*) => SwitchView())

; Right-click menu for File List
mnu_FileContext := Menu()
mnu_FileContext.Add(Txt.AddFile, (*) => AddCustomFile()) ; Always show Add option
mnu_FileContext.Add(Txt.RemoveFile, (*) => RemoveCustomFile()) ; Show only if an item is selected, otherwise show Add and Clear options
mnu_FileContext.Add(Txt.ClearList, (*) => ClearFileList()) ; Show context menu on right-click

lv_Icons.OnEvent("ContextMenu", ShowContextMenu)
lv_Icons.OnEvent("DoubleClick", CopyCurrentCode)
lv_Icons.OnEvent("ItemSelect", ShowPreview)

lst_Files.OnEvent("ContextMenu", ShowFileContextMenu)

; ========== RIGHT PANEL ==========
lbl_Preview := mGui.AddText("x680 y10 w210", "🔍 Preview:") ; Preview label
lbl_ExportSize := mGui.AddText("x680 y10 w40", "Size:") ; Size label
ddl_ExportSize := mGui.AddDropDownList("x680 y8 w70", ["16", "24", "32", "48", "64", "128", "256"]) ; Export size drop-down list
ddl_ExportSize.Text := ExportIconSize
ddl_ExportSize.OnEvent("Change", OnExportSizeChange)
lbl_IconNo := mGui.AddText("x680 y35 w256", "") ; Icon number text
pic_Preview := mGui.AddPicture("x680 y55 w128 h128 Border") ; Icon preview picture

btn_CopyImage := mGui.AddButton("x680 y190 w100 h30", Txt.Copy) ; Copy button
btn_CopyImage.OnEvent("Click", (*) => CopyIconToClipboard())

btn_SaveImage := mGui.AddButton("x+5 y190 w100 h30", Txt.Save) ; Save button
btn_SaveImage.OnEvent("Click", (*) => SaveIconToFile())

edt_IconCode := mGui.AddEdit("x680 y230 w210 h50", "") ; Icon code edit box

btn_CopyCode := mGui.AddButton("x680 y285 w95 h30", Txt.CopyCode) ; Copy code button
btn_CopyCode.OnEvent("Click", (*) => CopyCurrentCode())

btn_Test := mGui.AddButton("x+5 y285 w95 h30", Txt.Test) ; Test button
btn_Test.OnEvent("Click", (*) => TestIcon())

lbl_Favorites := mGui.AddText("x680 y330 w210", "⭐ Favorites:")
lv_Favorites := mGui.AddListView("x680 y350 w210 h150 Grid -Multi", ["Icon", "Num", "File"]) ; Favorites ListView
lv_Favorites.ModifyCol(2, "Integer"), lv_Favorites.ModifyCol(3, 100)
lv_Favorites.OnEvent("ItemSelect", ShowFavoritePreview)

btn_FavAdd := mGui.AddButton("x680 y505 w65 h30", Txt.Add) ; Add to favorites button
btn_FavAdd.OnEvent("Click", AddToFavorites)
btn_FavRemove := mGui.AddButton("x+5 y505 w65 h30", Txt.Remove) ; Remove from favorites button
btn_FavRemove.OnEvent("Click", RemoveFromFavorites)
btn_FavClear := mGui.AddButton("x+5 y505 w65 h30", Txt.Clear) ; Clear favorites button
btn_FavClear.OnEvent("Click", ClearFavorites)

; ========== BOTTOM PANEL ==========
lbl_Status := mGui.AddText("x10 y545 w880", "💡 Select a file ")

; Create ImageLists (initially)
IL_Small := IL_Create(100, 5, false)
IL_Large := IL_Create(100, 5, true)
IL_Favorites := IL_Create(100, 5, false)

; Default DLL files list
DefaultDllFiles := [
	"imageres.dll",
	"shell32.dll",
	"user32.dll",
	"ddores.dll",
	"ieframe.dll",
	"mmcndmgr.dll",
	"moricons.dll",
	"netcenter.dll",
	"netshell.dll",
	"networkexplorer.dll",
	"pifmgr.dll",
	"pnidui.dll",
	"setupapi.dll",
	"wmploc.dll",
	"wpdshext.dll",
	"compstui.dll",
	"accessibilitycpl.dll"
]

; Fill DLL list
InitializeDllList() ; Load saved files or defaults into the left panel list
LoadFavorites() ; Load favorites from file
Gdip.Startup() ; Initialize GDI+
OnMessage(0x0100, OnLvKeyDown) ; Handle arrow keys for linear navigation in ListView

; Event handlers
lst_Files.OnEvent("Change", LoadIcons)

; Show GUI
mGui.OnEvent("Close", (*) => CloseApplication())
mGui.OnEvent("Size", GuiSize)
mGui.OnEvent("DropFiles", OnDropFiles)
mGui.Show("w950 h670")
return

; ========== FUNCTIONS ==========

; Initialize DLL list with default files
InitializeDllList() {
	global lst_Files, DefaultDllFiles, SettingsFile, dllFileListChanged
	; Load saved files from settings
	if FileExist(SettingsFile) {
		try {
			savedFiles := FileRead(SettingsFile)
			loop parse, savedFiles, "`n", "`r" {
				if (A_LoopField != "" && FileExist(A_LoopField))
					lst_Files.Add([A_LoopField])
			}
		}
	}
	; If no saved files, load defaults
	if (SendMessage(0x018B, 0, 0, lst_Files.Hwnd) == 0) {
		for dllName in DefaultDllFiles {
			dllPath := A_WinDir "\System32\" dllName
			if FileExist(dllPath)
				lst_Files.Add([dllPath])
		}
		dllFileListChanged := true
	}
}

; Load icons from the selected DLL/EXE/ICO file
LoadIcons(*) {
	global lv_Icons, lst_Files, lbl_Status, CurrentDllPath, CurrentViewMode
	global IL_Small, IL_Large

	; Check if a file is selected in the list
	selectedIndex := lst_Files.Value
	if (selectedIndex = 0)
		return

	; Avoid reloading if the same file is already active
	selectedFile := lst_Files.Text
	if (selectedFile = CurrentDllPath)
		return

	CurrentDllPath := selectedFile

	; Clear existing UI data and icon arrays
	lv_Icons.Delete()

	; Destroy previous ImageLists to free up memory
	try IL_Destroy(IL_Small)
	try IL_Destroy(IL_Large)

	; Create new ImageLists (Initial capacity: 100, Grow: 5)
	IL_Small := IL_Create(100, 5, false)
	IL_Large := IL_Create(100, 5, true)

	lbl_Status.Text := "⏳ Loading icons: " . CurrentDllPath . " "

	iconCountTotal := DllCall("User32.dll\PrivateExtractIconsW",
		"Str", CurrentDllPath,
		"Int", 0, "Int", 0, "Int", 0,
		"Ptr", 0, "Ptr", 0,
		"UInt", 0, "UInt", 0, "UInt")

	; If no icons are found, update status and exit
	if (iconCountTotal <= 0) {
		lbl_Status.Text := "❌ No icons found in: " . CurrentDllPath
		return
	}

	iconLoadedCount := 0

	lv_Icons.Opt("-Redraw") ; Disable redraw during loading for performance
	; Loop through the total number of icons found
	loop iconCountTotal {
		iconIndex := A_Index

		; Add icon to small ImageList
		hSmall := IL_Add(IL_Small, CurrentDllPath, iconIndex)
		if (hSmall = 0)
			continue

		; Add icon to large ImageList
		hLarge := IL_Add(IL_Large, CurrentDllPath, iconIndex)
		if (hLarge = 0)
			continue

		iconLoadedCount++
		lv_Icons.Add("Icon" . hSmall, iconIndex) ; Add directly to ListView

		; Update status bar progress every 20 icons for better UI responsiveness
		if (Mod(iconLoadedCount, 20) = 0) {
			lbl_Status.Text := "⏳ Loaded: " . iconLoadedCount . " / " . iconCountTotal . " icons "
		}
	}
	lv_Icons.Opt("+Redraw") ; Re-enable redraw

	; Apply current view mode
	isReportView := CurrentViewMode < 2
	isLargeIcon := (CurrentViewMode = 1) || (CurrentViewMode = 3)
	lv_Icons.Opt(isReportView ? "+Report" : "+Icon")
	lv_Icons.SetImageList(isLargeIcon ? IL_Large : IL_Small, isReportView)

	; Final status update
	lbl_Status.Text := "✅ " . iconLoadedCount . " icons loaded | File: " . CurrentDllPath
}

; Show preview of selected icon
ShowPreview(*) {
	global lv_Icons, CurrentDllPath

	selectedRow := lv_Icons.GetNext(0, "Focused")
	if (selectedRow = 0) {
		UpdatePreviewPane("", 0)
		return
	}

	iconNum := GetIconNumberFromRow(selectedRow)
	UpdatePreviewPane(CurrentDllPath, iconNum)
}

; Update the entire preview pane (image, labels, code) based on an icon
UpdatePreviewPane(iconPath, iconNum) {
	global pic_Preview, lbl_IconNo, edt_IconCode, PreviewIcon, ExportIconSize

	if (iconPath = "" || iconNum = 0) {
		pic_Preview.Value := ""
		lbl_IconNo.Text := ""
		edt_IconCode.Value := ""
		PreviewIcon := { Path: "", Num: 0 }
		return
	}

	; Update global state
	PreviewIcon := { Path: iconPath, Num: iconNum }

	; Update GUI
	try pic_Preview.Value := "*Icon" iconNum " *w" ExportIconSize " *h" ExportIconSize " " iconPath
	catch {
		pic_Preview.Value := ""
	}

	lbl_IconNo.Text := "Icon #" iconNum
	edt_IconCode.Value := 'TraySetIcon("' iconPath '", ' iconNum ')'
}

; Switch view mode
SwitchView(*) {
	global lv_Icons, IL_Small, IL_Large, CurrentViewMode, mGui

	CurrentViewMode := Mod(CurrentViewMode + 1, 4)
	/*
	CurrentViewMode Values:
	0: Small Icon - Report View (Small Report)
	1: Large Icon - Report View (Large Report)
	2: Small Icon - Icon View (Small Icon)
	3: Large Icon - Icon View (Large Icon)
	*/

	isReportView := CurrentViewMode < 2
	isLargeIcon := (CurrentViewMode = 1) || (CurrentViewMode = 3)

	lv_Icons.Opt(isReportView ? "+Report" : "+Icon")
	lv_Icons.SetImageList(isLargeIcon ? IL_Large : IL_Small, isReportView)
}

; Get icon number from ListView row
GetIconNumberFromRow(rowNumber) {
	global lv_Icons
	if (rowNumber = 0)
		return 0

	; Get the text from the first column, which is the icon number.
	text := lv_Icons.GetText(rowNumber, 1)
	return Integer(text)
}

; Copy current icon code to clipboard
CopyCurrentCode(*) {
	global edt_IconCode, lbl_Status

	code := edt_IconCode.Value
	if (code = "") {
		ShowTempTooltip("⚠️ Select an icon first!")
		return
	}

	A_Clipboard := code
	lbl_Status.Text := "📋 Copied: " code
	ShowTempTooltip("📋 Code copied!")
}

; copies picture as png to clipboard with alpha-channel
CopyIconToClipboard() {
	global lbl_Status, PreviewIcon, ExportIconSize

	if (PreviewIcon.Path = "") {
		ShowTempTooltip("⚠️ Select an icon first!")
		return
	}

	try {
		iconNum := PreviewIcon.Num - 1
		hIcon := 0

		DllCall("PrivateExtractIcons", "Str", PreviewIcon.Path, "Int", iconNum,
			"Int", ExportIconSize, "Int", ExportIconSize, "Ptr*", &hIcon, "Ptr", 0, "UInt", 1, "UInt", 0)

		if (!hIcon)
			throw Error("Failed to extract icon")

		pBitmap := 0
		DllCall("gdiplus\GdipCreateBitmapFromHICON", "Ptr", hIcon, "Ptr*", &pBitmap)

		clsid := Buffer(16)
		DllCall("ole32\CLSIDFromString", "WStr", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "Ptr", clsid)

		DllCall("ole32\CreateStreamOnHGlobal", "Ptr", 0, "Int", true, "Ptr*", &pStream := 0)
		DllCall("gdiplus\GdipSaveImageToStream", "Ptr", pBitmap, "Ptr", pStream, "Ptr", clsid, "Ptr", 0)

		DllCall("ole32\GetHGlobalFromStream", "Ptr", pStream, "Ptr*", &hMem := 0)

		fmtPNG := DllCall("RegisterClipboardFormat", "Str", "PNG", "UInt")
		DllCall("OpenClipboard", "Ptr", 0)
		DllCall("EmptyClipboard")
		DllCall("SetClipboardData", "UInt", fmtPNG, "Ptr", hMem)
		DllCall("CloseClipboard")
	}
	finally {
		if (pBitmap)
			DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
		if (hIcon)
			DllCall("DestroyIcon", "Ptr", hIcon)

	}

	lbl_Status.Text := "📋 Icon copied to clipboard!"
	ShowTempTooltip("📋 Icon copied!")
}

; Save current icon to file
SaveIconToFile(*) {
	global lbl_Status, PreviewIcon, mGui, ExportIconSize

	if (PreviewIcon.Path = "") {
		ShowTempTooltip("⚠️ Select an icon first!")
		return
	}

	iconPath := PreviewIcon.Path
	iconNum := PreviewIcon.Num

	SplitPath(iconPath, &fileName)
	safeFileName := StrReplace(fileName, ".", "_")
	defaultSaveName := safeFileName "_Icon_" iconNum

	; SaveFile(Owner, Title, Filter, DefaultFileName, Options)
	saved := SaveFile([mGui.Hwnd, "Save Icon"]
		, defaultSaveName
		, { PNG: "*.png`n", BMP: "*.bmp", JPEG: "*.jpg", ICO: "*.ico" }
		, ""
		, 0x6)

	if !saved
		return
	savePath := saved.FileFullPath

	try {
		; Extract icon as HICON with alpha via PrivateExtractIcons
		iconIndex := iconNum - 1 ; PrivateExtractIcons is zero-based
		hIcon := 0
		DllCall("PrivateExtractIcons", "Str", iconPath, "Int", iconIndex,
			"Int", ExportIconSize, "Int", ExportIconSize, "Ptr*", &hIcon, "Ptr", 0, "UInt", 1, "UInt", 0)

		if (!hIcon)
			throw Error("Failed to extract icon")

		; Save via GDI+ (keeps alpha)
		Gdip.SaveHICONToFile(hIcon, savePath)

		lbl_Status.Text := "✅ Saved: " savePath
		ShowTempTooltip("✅ Icon saved!")
	} catch as err {
		MsgBox("Error saving image: " err.Message, "Error", "Icon!")
	} finally {
		if (hIcon)
			DllCall("DestroyIcon", "Ptr", hIcon)
	}
}
; Show context menu on right-click
ShowContextMenu(LV, Item, IsRightClick, X, Y) {
	if !IsRightClick ; No right-click
		return

	if (Item > 0) {
		LV.Modify(Item, "Select Focus Vis") ; Select and focus the item
		mnu_IconContext.Show(X, Y) ; Show the context menu
	} else {
		mnu_IconEmptyContext.Show(X, Y)
	}
}

; Show context menu for File List
ShowFileContextMenu(GuiCtrlObj, Item, IsRightClick, X, Y) {
	if (Item > 0)
		GuiCtrlObj.Choose(Item) ; Select the right-clicked item
	mnu_FileContext.Show(X, Y)
}

; Test icon by setting it as application icon
TestIcon(*) {
	global lv_Icons, CurrentDllPath, mGui, PreviewIcon

	testPath := ""
	testNum := 0

	if (IsSet(PreviewIcon) && PreviewIcon.Path != "") {
		testPath := PreviewIcon.Path
		testNum := PreviewIcon.Num
	} else if (lv_Icons.GetNext(0, "Focused") > 0) {
		testNum := GetIconNumberFromRow(lv_Icons.GetNext(0, "Focused"))
		testPath := CurrentDllPath
	}

	if (testPath == "") {
		MsgBox("Please select an icon from the list to test.", "Warning", "Icon!")
		return
	}

	try {
		; Change tray icon
		TraySetIcon(testPath, testNum)

		; Change window icon (WM_SETICON)
		if (hIconSmall := LoadPicture(testPath, "Icon" . testNum . " w16 h16", &isIcon))
			SendMessage(0x80, 0, hIconSmall, mGui.Hwnd) ; ICON_SMALL
		if (hIconBig := LoadPicture(testPath, "Icon" . testNum . " w32 h32", &isIcon))
			SendMessage(0x80, 1, hIconBig, mGui.Hwnd) ; ICON_BIG

		ShowTempTooltip("✅ Application icon updated", 2000)
	} catch as err {
		MsgBox("Error changing icon: " err.Message, "Error", "Icon!")
	}
}

; Add current icon to favorites
AddToFavorites(*) {
	global lv_Icons, CurrentDllPath, lbl_Status, lv_Favorites, IL_Favorites, FavoriteIcons, FavoritesListChanged

	selectedRow := lv_Icons.GetNext(0, "Focused")
	if (selectedRow = 0) {
		ShowTempTooltip("⚠️ Select an icon first!")
		return
	}

	iconNum := GetIconNumberFromRow(selectedRow)
	filePath := CurrentDllPath

	for fav in FavoriteIcons {
		if (fav.path = filePath && fav.num = iconNum) {
			ShowTempTooltip("⚠️ Icon already in favorites!")
			return
		}
	}

	FavoriteIcons.Push({ path: filePath, num: iconNum })
	FavoritesListChanged := true

	iconIndex := IL_Add(IL_Favorites, filePath, iconNum)
	SplitPath(filePath, &fileName)
	lv_Favorites.Add("Icon" . iconIndex, "", iconNum, fileName)

	lbl_Status.Text := "⭐ Added to favorites: " fileName " " iconNum
	ShowTempTooltip("⭐ Added to favorites!")
}

; Remove selected icon from favorites
RemoveFromFavorites(*) {
	global lv_Favorites, FavoriteIcons, FavoritesListChanged

	selected := lv_Favorites.GetNext(0, "Focused")
	if (selected = 0) {
		ShowTempTooltip("⚠️ Select a favorite to remove!")
		return
	}

	FavoriteIcons.RemoveAt(selected)
	lv_Favorites.Delete(selected)
	FavoritesListChanged := true

	ShowTempTooltip("➖ Favorite removed.")
}

; Clear all favorites
ClearFavorites(*) {
	global lv_Favorites, FavoriteIcons, IL_Favorites, lbl_Status, FavoritesListChanged

	if (FavoriteIcons.Length = 0)
		return

	if (MsgBox("Are you sure you want to clear all favorites?", "Clear Favorites", "YesNo Icon?") = "No")
		return

	FavoriteIcons := []
	lv_Favorites.Delete()
	try IL_Destroy(IL_Favorites)
	IL_Favorites := IL_Create(100, 5, false)
	lv_Favorites.SetImageList(IL_Favorites)
	FavoritesListChanged := true

	lbl_Status.Text := "🗑️ Favorites cleared"
	ShowTempTooltip("🗑️ Favorites cleared")
}

; Load favorites from file and update menu
LoadFavorites() {
	global lv_Favorites, IL_Favorites, FavoritesFile, FavoriteIcons, FavoritesListChanged

	lv_Favorites.Delete()
	try IL_Destroy(IL_Favorites)
	IL_Favorites := IL_Create(100, 5, false)
	lv_Favorites.SetImageList(IL_Favorites)
	FavoriteIcons := []

	if FileExist(FavoritesFile) {
		loop read, FavoritesFile {
			favString := A_LoopReadLine
			if (Trim(favString) = "")
				continue

			parts := StrSplit(favString, "|")
			if (parts.Length = 2 && FileExist(parts[1])) {
				filePath := parts[1]
				iconNum := Integer(parts[2])

				FavoriteIcons.Push({ path: filePath, num: iconNum })

				iconIndex := IL_Add(IL_Favorites, filePath, iconNum)
				SplitPath(filePath, &fileName)
				lv_Favorites.Add("Icon" . iconIndex, "", iconNum, fileName)
			}
		}
	}
	FavoritesListChanged := false
}

SaveFavorites() {
	global FavoritesFile, FavoriteIcons, FavoritesListChanged

	if !FavoritesListChanged && FileExist(FavoritesFile)
		return

	fileContent := ""
	for fav in FavoriteIcons {
		fileContent .= fav.path "|" fav.num "`n"
	}
	if FileExist(FavoritesFile)
		FileDelete(FavoritesFile)
	FileAppend(fileContent, FavoritesFile)
	FavoritesListChanged := false
}

ShowFavoritePreview(*) {
	global lv_Favorites, FavoriteIcons

	selected := lv_Favorites.GetNext(0, "Focused")
	if (selected = 0) {
		UpdatePreviewPane("", 0)
		return
	}

	fav := FavoriteIcons[selected]
	UpdatePreviewPane(fav.path, fav.num)
}

; Add custom DLL/EXE/ICO file to the list
AddCustomFile(*) {
	global lst_Files, lbl_Status, dllFileListChanged

	selectedFile := FileSelect(3, , "Select Icon Source", "Source Files (*.dll; *.exe; *.ico; *.lnk)")

	if (selectedFile = "")
		return

	if !FileExist(selectedFile) {
		MsgBox("File not found!", "Error", "Icon!")
		return
	}

	; If file is .lnk, get target path
	filePathToAdd := selectedFile
	SplitPath(selectedFile, , , &ext)
	if (ext = "lnk") {
		try {
			FileGetShortcut(selectedFile, &targetPath)
			filePathToAdd := targetPath
		} catch {
			MsgBox("Could not resolve shortcut: " . selectedFile, "Error", "Icon!")
			return
		}
	}

	; Check if already in list
	if (IsFileInList(filePathToAdd)) {
		MsgBox("This file is already in the list!", "Info", "Iconi")
		return
	}

	; Add to list
	lst_Files.Add([filePathToAdd])
	dllFileListChanged := true
	newCount := SendMessage(0x018B, 0, 0, lst_Files.Hwnd) ; LB_GETCOUNT
	lst_Files.Choose(newCount)
	LoadIcons()
	lbl_Status.Text := "✅ File added"
}

; Remove selected file from list
RemoveCustomFile(*) {
	global lst_Files, dllFileListChanged

	selected := lst_Files.Value
	if (selected = 0) {
		MsgBox("Please select a file to remove.", "Info", "Icon!")
		return
	}

	lst_Files.Delete(selected)
	dllFileListChanged := true
}

; Clear all files from list
ClearFileList(*) {
	global lst_Files, lv_Icons, CurrentDllPath, lbl_Status, dllFileListChanged

	if (SendMessage(0x018B, 0, 0, lst_Files.Hwnd) == 0) ; LB_GETCOUNT
		return

	if (MsgBox("Are you sure you want to clear the file list?", "Clear List", "YesNo Icon?") = "No")
		return

	lst_Files.Delete()
	lv_Icons.Delete()
	CurrentDllPath := ""
	dllFileListChanged := true
	lbl_Status.Text := "🗑️ List cleared"
}

; Save file list and exit
CloseApplication(*) {
	SaveFileList() ; Save file list to disk
	SaveFavorites() ; Save favorites to disk
	Gdip.Shutdown() ; Cleanup GDI+ resources
	ExitApp() ; Close application
}

; Save current file list to disk
SaveFileList() {
	global lst_Files, SettingsFile, dllFileListChanged

	if !dllFileListChanged && FileExist(SettingsFile)
		return

	try {
		if FileExist(SettingsFile)
			FileDelete(SettingsFile)

		items := ControlGetItems(lst_Files)
		f := FileOpen(SettingsFile, "w")
		for item in items {
			f.WriteLine(item)
		}
		f.Close()
		dllFileListChanged := false
	}
}

; Check if a file path already exists in the file list
IsFileInList(filePath) {
	global lst_Files
	for item in ControlGetItems(lst_Files) {
		if (item = filePath)
			return true
	}
	return false
}

; Handle files dropped onto the GUI
OnDropFiles(GuiObj, GuiCtrlObj, FileArray, X, Y) {
	global lst_Files, lbl_Status, dllFileListChanged

	addedCount := 0

	for droppedPath in FileArray {
		filePathToAdd := droppedPath
		SplitPath(filePathToAdd, , , &ext)

		if (ext = "lnk") {
			try {
				FileGetShortcut(filePathToAdd, &targetPath)
				filePathToAdd := targetPath
			} catch {
				continue ; Skip if shortcut cannot be resolved
			}
		}

		SplitPath(filePathToAdd, , , &ext) ; Get target file extension
		if (ext != "dll" && ext != "exe" && ext != "ico")
			continue

		; Already in the list?
		if !IsFileInList(filePathToAdd) {
			lst_Files.Add([filePathToAdd])
			addedCount++
			dllFileListChanged := true
		}
	}

	if (addedCount > 0) {
		lbl_Status.Text := "✅ " addedCount " file(s) added"
		newCount := SendMessage(0x018B, 0, 0, lst_Files.Hwnd) ; LB_GETCOUNT
		lst_Files.Choose(newCount)
		LoadIcons()
	} else {
		lbl_Status.Text := "ℹ️ No file added (already in list or not supported)"
	}
}

; Handle GUI resizing
GuiSize(GuiObj, MinMax, Width, Height) {
	if (MinMax = -1)
		return

	global mGui, lst_Files, lv_Icons, lbl_Status, CurrentViewMode, pic_Preview, lbl_IconNo, edt_IconCode, lbl_Preview,
		btn_CopyImage, btn_SaveImage, btn_CopyCode, btn_Test, btn_Switch, lv_Favorites, btn_FavAdd, btn_FavRemove,
		lbl_Favorites, btn_FavClear, lbl_ExportSize, ddl_ExportSize, ExportIconSize, LastGuiW, LastGuiH

	LastGuiW := Width
	LastGuiH := Height

	GuiRedraw(GuiObj, 0)

	; Left panel
	lst_Files.Move(, , , Height - 110)

	; Right Panel
	rightPanelWidth := 210
	if (ExportIconSize + 20 > rightPanelWidth)
		rightPanelWidth := ExportIconSize + 20
	rightPanelX := Width - rightPanelWidth - 10

	sizeComboW := 70
	sizeLabelW := 35
	sizeComboX := rightPanelX + rightPanelWidth - sizeComboW
	sizeLabelX := sizeComboX - sizeLabelW - 5

	lbl_Preview.Move(rightPanelX, 10, sizeLabelX - rightPanelX - 5)
	lbl_ExportSize.Move(sizeLabelX, 10, sizeLabelW)
	ddl_ExportSize.Move(sizeComboX, 8, sizeComboW)

	lbl_IconNo.Move(rightPanelX, 35)
	pic_Preview.Move(rightPanelX, 55, ExportIconSize, ExportIconSize)

	previewBottom := 55 + ExportIconSize
	buttonsY := previewBottom + 7
	btn_CopyImage.Move(rightPanelX, buttonsY)
	btn_SaveImage.Move(rightPanelX + 105, buttonsY)

	iconCodeY := buttonsY + 40
	edt_IconCode.Move(rightPanelX, iconCodeY, rightPanelWidth, 50)

	copyCodeY := iconCodeY + 55
	btn_CopyCode.Move(rightPanelX, copyCodeY)
	btn_Test.Move(rightPanelX + 105, copyCodeY)

	favoritesLabelY := copyCodeY + 45
	lbl_Favorites.Move(rightPanelX, favoritesLabelY)

	favButtonsY := Height - 70
	favoritesListY := favoritesLabelY + 20
	favListHeight := favButtonsY - favoritesListY - 5
	if (favListHeight < 0)
		favListHeight := 0
	lv_Favorites.Move(rightPanelX, favoritesListY, rightPanelWidth, favListHeight)
	lv_Favorites.ModifyCol(3, rightPanelWidth - 38 - 50 - 10)

	btnWidth := 65
	btn_FavAdd.Move(rightPanelX, favButtonsY, btnWidth)
	btn_FavRemove.Move(rightPanelX + btnWidth + 5, favButtonsY, btnWidth)
	btn_FavClear.Move(rightPanelX + (btnWidth + 5) * 2, favButtonsY, btnWidth)

	; Middle Panel
	middlePanelStartX := 270 ; Starting X coordinate of IconListView
	middlePanelGapToRight := 10 ; Gap between the middle and right panels
	newListViewWidth := rightPanelX - middlePanelStartX - middlePanelGapToRight
	lv_Icons.Move(middlePanelStartX, 35, newListViewWidth, Height - 75)

	isReportView := CurrentViewMode < 2
	if (isReportView)
		lv_Icons.ModifyCol(1, newListViewWidth - 20)

	btn_Switch.Move(middlePanelStartX + newListViewWidth - 140, 5)

	; Bottom status bar
	lbl_Status.Move(, Height - 30, Width - 20)

	; Force rearrange icons if in Icon view
	if (CurrentViewMode >= 2)
		SendMessage(0x1016, 0, 0, lv_Icons.Hwnd) ; LVM_ARRANGE

	GuiRedraw(GuiObj, 1) ; Enable redraw
	WinRedraw(GuiObj) ; Force redraw to clear artifacts
}

; Enable or disable GUI redraw
GuiRedraw(GuiObj, redraw) {
	SendMessage(0x000B, redraw, 0, GuiObj.Hwnd) ; WM_SETREDRAW = 1 (On)
}

; ========== HELPER FUNCTIONS ==========

; Show a tooltip for a short duration
ShowTempTooltip(message, duration := 1500) {
	ToolTip(message, , , 1)
	SetTimer(() => ToolTip(, , , 1), -duration)
}

OnExportSizeChange(*) {
	global ddl_ExportSize, ExportIconSize, LastGuiW, LastGuiH, mGui, PreviewIcon

	size := Integer(ddl_ExportSize.Text)
	if (size < 16)
		size := 16
	ExportIconSize := size

	if (LastGuiW > 0 && LastGuiH > 0)
		GuiSize(mGui, 0, LastGuiW, LastGuiH)

	if (PreviewIcon.Path != "")
		UpdatePreviewPane(PreviewIcon.Path, PreviewIcon.Num)
}

; Handle arrow key navigation in ListView
OnLvKeyDown(wParam, lParam, msg, hwnd) {
	global lv_Icons
	if (!IsSet(lv_Icons) || !lv_Icons) ; Safety check if GUI is not destroyed but variable is cleared
		return

	if (hwnd = lv_Icons.Hwnd) { ; Left or Right arrow keys
		dir := (wParam = 39) ? 1 : (wParam = 37) ? -1 : 0 ; Right = 39, Left = 37
		if (dir) { ; Arrow key pressed
			item := lv_Icons.GetNext(0, "Focused") ; Get currently focused item
			nextItem := item + dir ; Calculate next item index
			if (nextItem > 0 && nextItem <= lv_Icons.GetCount()) { ; Valid next item
				lv_Icons.Modify(0, "-Focus -Select") ; Remove focus and selection from current item
				lv_Icons.Modify(nextItem, "+Focus +Select +Vis") ; Focus, select, and ensure next item is visible
				ShowPreview() ; Update preview for new selection
				return 0 ; Prevent default handling
			}
		}
	}
}

; Show About dialog
ShowAbout(*) {
	aboutText := A_ScriptName
	aboutText .= "
(
`n`nA tool for viewing and extracting icons from DLL, EXE, and ICO files.`n
Mesut Akcan
mesutakcan.blogspot.com
github.com/akcansoft
youtube.com/mesutakcan
)"

	MsgBox(aboutText, "About", "Iconi")
}
