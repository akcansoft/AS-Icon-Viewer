;@Ahk2Exe-SetName AS Icon Viewer
;@Ahk2Exe-SetDescription AS Icon Viewer
;@Ahk2Exe-SetFileVersion 1.3
;@Ahk2Exe-SetCompanyName AkcanSoft
;@Ahk2Exe-SetCopyright ©2026 Mesut Akcan
;@Ahk2Exe-SetMainIcon app_icon.ico
;@Ahk2Exe-ExeName "AS Icon Viewer.exe"

; AS Icon Viewer
; 10/02/2026

; Mesut Akcan
; -----------
; mesutakcan.blogspot.com
; github.com/akcansoft
; youtube.com/mesutakcan

#Requires AutoHotkey v2
#SingleInstance Force

#Include Gdip.ahk ; GDI+
#Include SaveFileDialog.ahk ; Custom Save File Dialog

; ========== TRAY ICON ==========
; Set tray icon for source code (ignored in compiled EXE)
;@Ahk2Exe-IgnoreBegin
try TraySetIcon(A_ScriptDir "\app_icon.ico")
;@Ahk2Exe-IgnoreEnd

; Global variables
A_ScriptName := "AS Icon Viewer v1.3"
global SettingsFile := A_ScriptDir "\saved_files.txt" ; For storing user-added files in the left panel
global FavoritesFile := A_ScriptDir "\favorites.txt" ; For storing favorite icons
global CurrentDllPath := "" ; Currently loaded DLL/EXE/ICO file path
global CurrentViewMode := 3 ; 0:SmallReport, 1:LargeReport, 2:SmallIcon, 3:LargeIcon
global IL_Small := 0 ; Small ImageList
global IL_Large := 0 ; Large ImageList
global lv_Favorites := 0 ; Favorites ListView
global IL_Favorites := 0 ; Favorites ImageList
global dllFileListChanged := false ; To track changes in the DLL file list for saving
global FavoritesListChanged := false ; To track changes in the favorites list for saving
global PreviewIcon := { Path: "", Num: 0 } ; Currently previewed icon info
global ExportIconSize := 128 ; Default export size
global LastGuiW := 0 ; For tracking GUI size changes
global LastGuiH := 0

global Symbol := {
	Info: "ℹ️",
	Success: "✅",
	Warning: "⚠️",
	Error: "❌",
	Copy: "📋",
	Save: "💾",
	Star: "⭐",
	Remove: "➖",
	Add: "➕",
	Clear: "🧹",
	Test: "🧪",
	File: "📂",
	Trash: "🗑️",
	Loading: "⏳",
	Color: "🎨",
	Search: "🔍",
	About: "❓",
	Web: "🌐"
}

global Txt := {
	File: "&File",
	View: "Icon &View",
	Favorites: "&Favorites",
	Icon: "&Icon",
	Help: "&Help",
	Add: Symbol.Add " &Add",
	Remove: Symbol.Remove " &Remove",
	Clear: Symbol.Clear " C&lear",
	AddFav: Symbol.Add " A&dd",
	RemoveFav: Symbol.Remove " Rem&ove",
	ClearFav: Symbol.Clear " Clea&r",
	Save: Symbol.Save " &Save",
	Test: Symbol.Test " &Test icon",
	Copy: Symbol.Copy " &Copy",
	CopyCode: Symbol.Copy " C&opy Code",
	AddFile: Symbol.Add " &Add File ",
	RemoveFile: Symbol.Remove " &Remove File",
	ClearList: Symbol.Clear " &Clear List",
	OpenFolder: Symbol.File " &Open File Location",
	Exit: "⏻ E&xit",
	Refresh: "&Refresh",
	CopyImage: Symbol.Copy " Co&py Image",
	SaveImage: Symbol.Save " &Save Image ",
	AddToFavorites: Symbol.Add " &Add to Favorites",
	RemoveFromFavorites: Symbol.Remove " &Remove from Favorites",
	About: Symbol.About " &About",
	Website: Symbol.Web " &Visit Website",
	GitHub: Symbol.Web " &GitHub Repository",
	View0: "&Small Report",
	View1: "&Large Report",
	View2: "S&mall Icon",
	View3: "L&arge Icon"
}

; ========== TRAY MENU ==========
Tray := A_TrayMenu
Tray.Delete() ; Delete the standard items.
Tray.Add "Open " A_ScriptName, (*) => mGui.Show()
Tray.Add() ; separator
Tray.AddStandard()
Tray.Default := "Open " A_ScriptName

; Create main GUI
mGui := Gui("+Resize +MinSize900x600", A_ScriptName)
mGui.SetFont("s9", "Segoe UI")

; ========== MENU BAR ==========
; File Menu
mnu_File := Menu()
mnu_File.Add(Txt.AddFile, (*) => AddCustomFile())
mnu_File.Add(Txt.RemoveFile, (*) => RemoveCustomFile())
mnu_File.Add(Txt.ClearList, (*) => ClearFileList())
mnu_File.Add() ; Separator
mnu_File.Add(Txt.Exit, (*) => CloseApplication())

; Icon View Menu
mnu_View := Menu()
mnu_View.Add(Txt.View0, (*) => SetViewMode(0))
mnu_View.Add(Txt.View1, (*) => SetViewMode(1))
mnu_View.Add(Txt.View2, (*) => SetViewMode(2))
mnu_View.Add(Txt.View3, (*) => SetViewMode(3))
mnu_View.Add() ; Separator
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
mnu_Help.Add(Txt.About, (*) => ShowAbout())
mnu_Help.Add() ; Separator
mnu_Help.Add(Txt.Website, (*) => Run("https://mesutakcan.blogspot.com"))
mnu_Help.Add(Txt.GitHub, (*) => Run("https://github.com/akcansoft/AS-Icon-Viewer"))

; Main Menu Bar
mnu_Main := MenuBar()
mnu_Main.Add(Txt.File, mnu_File)
mnu_Main.Add(Txt.View, mnu_View)
mnu_Main.Add(Txt.Favorites, mnu_Favorites)
mnu_Main.Add(Txt.Icon, mnu_Icon)
mnu_Main.Add(Txt.Help, mnu_Help)
mGui.MenuBar := mnu_Main

; ========== LEFT PANEL ==========
mGui.AddText("x10 y10 w250", Symbol.File " Icon Files:")
btn_AddFile := mGui.AddButton("x10 y35 w80 h30", Txt.Add) ; Add button
btn_AddFile.OnEvent("Click", AddCustomFile)
btn_RemoveFile := mGui.AddButton("x95 y35 w80 h30", Txt.Remove) ; Remove button
btn_RemoveFile.OnEvent("Click", RemoveCustomFile)
btn_ClearList := mGui.AddButton("x180 y35 w80 h30", Txt.Clear) ; Clear List button
btn_ClearList.OnEvent("Click", ClearFileList)
lst_Files := mGui.AddListBox("x10 y70 w250 h465 Multi +VScroll +HScroll", []) ; List of DLL/EXE/ICO files

; ========== MIDDLE PANEL ==========
mGui.AddText("x270 y10", Symbol.Color " Icons:")
ddl_ViewMode := mGui.AddDropDownList("x335 y5", [StrReplace(Txt.View0, "&"), StrReplace(Txt.View1, "&"), StrReplace(Txt
	.View2, "&"), StrReplace(Txt.View3, "&")])
ddl_ViewMode.OnEvent("Change", OnViewChange)

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

; Right-click menu for Icons (Empty space)
mnu_IconEmptyContext := Menu()
mnu_IconEmptyContext.Add(Txt.Refresh, (*) => LoadIcons())

; Icons ListView Event Handlers
lv_Icons.OnEvent("ContextMenu", ShowContextMenu)
lv_Icons.OnEvent("DoubleClick", CopyCurrentCode)
lv_Icons.OnEvent("ItemSelect", ShowPreview)

; File List Event Handlers
lst_Files.OnEvent("ContextMenu", ShowFileContextMenu)

; ========== RIGHT PANEL ==========
lbl_PreviewSize := mGui.AddText("x680 y10", Symbol.Search " Preview Size:") ; Preview Size label
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

lbl_Favorites := mGui.AddText("x680 y335 w210", Symbol.Star " Favorites:")
btn_FavAdd := mGui.AddButton("x680 y360 w65 h30", Txt.AddFav) ; Add to favorites button
btn_FavAdd.OnEvent("Click", AddToFavorites)
btn_FavRemove := mGui.AddButton("x+5 y360 w65 h30", Txt.RemoveFav) ; Remove from favorites button
btn_FavRemove.OnEvent("Click", RemoveFromFavorites)
btn_FavClear := mGui.AddButton("x+5 y360 w65 h30", Txt.ClearFav) ; Clear favorites button
btn_FavClear.OnEvent("Click", ClearFavorites)

lv_Favorites := mGui.AddListView("x680 y395 w210 h150 Grid", ["Icon", "Num", "File", "FullPath"]) ; Favorites ListView
lv_Favorites.ModifyCol(2, "Integer"), lv_Favorites.ModifyCol(3, 100), lv_Favorites.ModifyCol(4, 0) ; Hide 4th column
lv_Favorites.OnEvent("ItemSelect", ShowFavoritePreview)

; ========== BOTTOM PANEL ==========
sb_Status := mGui.AddStatusBar()
sb_Status.SetParts(250, 150) ; Divide status bar into 3 parts: Process | Icon Count | File Path
SetStatus("Select a file", Symbol.Info)

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
ApplyViewSettings() ; Set initial view state and menu checkmark

; Event handlers
lst_Files.OnEvent("Change", LoadIcons)

; Show GUI
mGui.OnEvent("Close", (*) => CloseApplication())
mGui.OnEvent("Size", GuiSize)
mGui.OnEvent("DropFiles", OnDropFiles)
mGui.Show("w950 h670")
return

; ========== FUNCTIONS ==========

OnViewChange(*) {
	global ddl_ViewMode
	SetViewMode(ddl_ViewMode.Value - 1)
}

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
	global lv_Icons, lst_Files, CurrentDllPath, CurrentViewMode
	global IL_Small, IL_Large

	; Check if a file is selected in the list
	selectedIndices := lst_Files.Value
	if (Type(selectedIndices) = "Array" ? selectedIndices.Length = 0 : selectedIndices = 0)
		return

	; Handle multi-select ListBox (Text returns an array)
	selectedFile := lst_Files.Text
	if (Type(selectedFile) = "Array")
		selectedFile := (selectedFile.Length > 0) ? selectedFile[1] : ""

	if (selectedFile = "" || selectedFile = CurrentDllPath)
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

	SetStatus("Loading icons: " . CurrentDllPath, Symbol.Loading)

	iconCountTotal := DllCall("User32.dll\PrivateExtractIconsW",
		"Str", CurrentDllPath,
		"Int", 0, "Int", 0, "Int", 0,
		"Ptr", 0, "Ptr", 0,
		"UInt", 0, "UInt", 0, "UInt")

	; If no icons are found, update status and exit
	if (iconCountTotal <= 0) {
		SetStatus("No icons found in: " . CurrentDllPath, Symbol.Error)
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

		if (Mod(iconLoadedCount, 20) = 0) {
			SetStatus("Loading icons...", Symbol.Loading, iconLoadedCount . " / " . iconCountTotal)
		}
	}
	lv_Icons.Opt("+Redraw") ; Re-enable redraw

	; Apply current view mode
	ApplyViewSettings()

	; Final status update
	SetStatus("Ready", Symbol.Success, iconLoadedCount . " icons loaded", CurrentDllPath)
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

; Set view mode
SetViewMode(mode) {
	global CurrentViewMode
	CurrentViewMode := mode
	ApplyViewSettings()
}

; Apply current view settings to ListView
ApplyViewSettings() {
	global lv_Icons, IL_Small, IL_Large, CurrentViewMode, mnu_View, mGui, LastGuiW, LastGuiH, Txt, ddl_ViewMode

	isReportView := CurrentViewMode < 2
	isLargeIcon := (CurrentViewMode = 1) || (CurrentViewMode = 3)

	lv_Icons.Opt(isReportView ? "+Report" : "+Icon")
	lv_Icons.SetImageList(isLargeIcon ? IL_Large : IL_Small, isReportView)

	; Update DDL
	ddl_ViewMode.Value := CurrentViewMode + 1

	; Update menu checkmarks
	try {
		mnu_View.Uncheck(Txt.View0), mnu_View.Uncheck(Txt.View1)
		mnu_View.Uncheck(Txt.View2), mnu_View.Uncheck(Txt.View3)

		target := (CurrentViewMode == 0) ? Txt.View0 : (CurrentViewMode == 1) ? Txt.View1 : (CurrentViewMode == 2) ?
			Txt.View2 : Txt.View3
		mnu_View.Check(target)
	}

	if (LastGuiW > 0 && LastGuiH > 0)
		GuiSize(mGui, 0, LastGuiW, LastGuiH)
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

; Copies the current icon code to the clipboard for use in scripts.
CopyCurrentCode(*) {
	global edt_IconCode

	code := edt_IconCode.Value
	if (code = "") {
		ShowTempTooltip("Select an icon first!", Symbol.Warning)
		return
	}

	A_Clipboard := code
	SetStatus("Copied: " code, Symbol.Copy)
	ShowTempTooltip("Code copied!", Symbol.Copy)
}

; Copies a picture as a PNG to the clipboard with its alpha channel preserved.
; Uses GDI+ to process the icon and convert it to a format the clipboard can handle for transparency.
CopyIconToClipboard() {
	global PreviewIcon, ExportIconSize

	if (PreviewIcon.Path = "") {
		ShowTempTooltip("Select an icon first!", Symbol.Warning)
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

	SetStatus("Icon copied to clipboard!", Symbol.Copy)
	ShowTempTooltip("Icon copied!", Symbol.Copy)
}

; Saves the currently selected icon to a file.
SaveIconToFile(*) {
	global PreviewIcon, mGui, ExportIconSize

	if (PreviewIcon.Path = "") {
		ShowTempTooltip("Select an icon first!", Symbol.Warning)
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

		SetStatus("Saved: " savePath, Symbol.Success)
		ShowTempTooltip("Icon saved!", Symbol.Success)
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
	if (!Item)
		return

	mnu_FileContext := Menu()
	if (Item > 0) {
		GuiCtrlObj.Choose(Item) ; Select the right-clicked item
		mnu_FileContext.Add(Txt.OpenFolder, (*) => OpenFileLocation())
		mnu_FileContext.Add(Txt.RemoveFile, (*) => RemoveCustomFile())
		mnu_FileContext.Add() ; Separator
	}
	mnu_FileContext.Add(Txt.AddFile, (*) => AddCustomFile())
	mnu_FileContext.Add(Txt.ClearList, (*) => ClearFileList())
	mnu_FileContext.Show(X, Y)
}

; Opens Windows Explorer with the selected file highlighted.
OpenFileLocation(*) {
	global lst_Files
	if (lst_Files.Value = 0)
		return

	filePath := lst_Files.Text
	if FileExist(filePath)
		Run('explorer.exe /select,"' filePath '"')
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

		ShowTempTooltip("Application icon updated", Symbol.Success, 2000)
	} catch as err {
		MsgBox("Error changing icon: " err.Message, "Error", "Icon!")
	}
}

; Add current icon to favorites
AddToFavorites(*) {
	global lv_Icons, CurrentDllPath, lv_Favorites, IL_Favorites, FavoritesListChanged

	selectedRow := lv_Icons.GetNext(0, "Focused")
	if (selectedRow = 0) {
		ShowTempTooltip("Select an icon first!", Symbol.Warning)
		return
	}

	iconNum := GetIconNumberFromRow(selectedRow)
	filePath := CurrentDllPath

	; Check for duplicates directly in ListView
	loop lv_Favorites.GetCount() {
		if (lv_Favorites.GetText(A_Index, 2) = String(iconNum) && lv_Favorites.GetText(A_Index, 4) = filePath) {
			ShowTempTooltip("Icon already in favorites!", Symbol.Warning)
			return
		}
	}

	FavoritesListChanged := true

	; Add to ListView with hidden full path
	iconIndex := IL_Add(IL_Favorites, filePath, iconNum)
	SplitPath(filePath, &fileName)
	lv_Favorites.Add("Icon" . iconIndex, "", iconNum, fileName, filePath) ; filePath to 4th column

	SetStatus("Added to favorites: " fileName " " iconNum, Symbol.Star)
	ShowTempTooltip("Added to favorites!", Symbol.Star)
}

; Remove selected icon from favorites
RemoveFromFavorites(*) {
	global lv_Favorites, FavoritesListChanged

	rowNumber := 0
	selectedRows := []
	loop {
		rowNumber := lv_Favorites.GetNext(rowNumber)
		if not rowNumber
			break
		selectedRows.Push(rowNumber)
	}

	if (selectedRows.Length = 0) {
		ShowTempTooltip("Select favorite(s) to remove!", Symbol.Warning)
		return
	}

	; Remove from bottom to top to preserve correct indices
	loop selectedRows.Length {
		idx := selectedRows[selectedRows.Length - A_Index + 1]
		lv_Favorites.Delete(idx)
	}

	FavoritesListChanged := true
	ShowTempTooltip("Favorite(s) removed.", Symbol.Remove)
}

; Clear all favorites
ClearFavorites(*) {
	global lv_Favorites, IL_Favorites, FavoritesListChanged

	if (lv_Favorites.GetCount() = 0)
		return

	if (MsgBox("Are you sure you want to clear all favorites?", "Clear Favorites", "YesNo Icon?") = "No")
		return

	lv_Favorites.Delete()
	try IL_Destroy(IL_Favorites)
	IL_Favorites := IL_Create(100, 5, false)
	lv_Favorites.SetImageList(IL_Favorites)
	FavoritesListChanged := true

	SetStatus("Favorites cleared", Symbol.Trash)
	ShowTempTooltip("Favorites cleared", Symbol.Trash)
}

; Load favorites from file and update menu
LoadFavorites() {
	global lv_Favorites, IL_Favorites, FavoritesFile, FavoritesListChanged

	lv_Favorites.Delete()
	try IL_Destroy(IL_Favorites)
	IL_Favorites := IL_Create(100, 5, false)
	lv_Favorites.SetImageList(IL_Favorites)

	if FileExist(FavoritesFile) {
		loop read, FavoritesFile {
			favString := A_LoopReadLine
			if (Trim(favString) = "")
				continue

			parts := StrSplit(favString, "|")
			if (parts.Length = 2 && FileExist(parts[1])) {
				filePath := parts[1]
				iconNum := Integer(parts[2])

				iconIndex := IL_Add(IL_Favorites, filePath, iconNum)
				SplitPath(filePath, &fileName)
				lv_Favorites.Add("Icon" . iconIndex, "", iconNum, fileName, filePath)
			}
		}
	}
	FavoritesListChanged := false
}

SaveFavorites() {
	global FavoritesFile, lv_Favorites, FavoritesListChanged

	if !FavoritesListChanged && FileExist(FavoritesFile)
		return

	fileContent := ""
	loop lv_Favorites.GetCount() {
		iconNum := lv_Favorites.GetText(A_Index, 2)
		filePath := lv_Favorites.GetText(A_Index, 4)
		fileContent .= filePath "|" iconNum "`n"
	}

	if FileExist(FavoritesFile)
		FileDelete(FavoritesFile)
	FileAppend(fileContent, FavoritesFile)
	FavoritesListChanged := false
}

ShowFavoritePreview(*) {
	global lv_Favorites

	selected := lv_Favorites.GetNext(0, "Focused")
	if (selected = 0) {
		UpdatePreviewPane("", 0)
		return
	}

	; Read directly from ListView columns to avoid sync issues with sorting
	iconNum := Integer(lv_Favorites.GetText(selected, 2))
	filePath := lv_Favorites.GetText(selected, 4)

	UpdatePreviewPane(filePath, iconNum)
}

; Add custom DLL/EXE/ICO file to the list
AddCustomFile(*) {
	selectedFiles := FileSelect("M3", , "Select Icon Source", "Source Files (*.dll; *.exe; *.ico; *.lnk)")
	if (Type(selectedFiles) = "Array")
		AddFilesToList(selectedFiles)
	else if (selectedFiles != "")
		AddFilesToList([selectedFiles])
}

; Processes a list of files to be added to the application.
; Handles shortcut resolution, duplicate checks, and extension validation.
AddFilesToList(FileArray) {
	global lst_Files, dllFileListChanged, Symbol

	addedCount := 0
	lastAddedPath := ""

	for file in FileArray {
		if (file = "" || !FileExist(file))
			continue

		filePathToAdd := file
		SplitPath(file, , , &ext)

		if (ext = "lnk") {
			try {
				FileGetShortcut(file, &targetPath)
				filePathToAdd := targetPath
			} catch {
				continue
			}
		}

		SplitPath(filePathToAdd, , , &ext)
		if !(ext ~= "i)^(dll|exe|ico)$")
			continue

		if (IsFileInList(filePathToAdd))
			continue

		lst_Files.Add([filePathToAdd])
		addedCount++
		lastAddedPath := filePathToAdd
		dllFileListChanged := true
	}

	if (addedCount > 0) {
		newCount := SendMessage(0x018B, 0, 0, lst_Files.Hwnd) ; LB_GETCOUNT
		lst_Files.Choose(newCount)
		LoadIcons()

		statusMsg := (addedCount = 1) ? "File added" : addedCount " files added"
		SetStatus(statusMsg, Symbol.Success, , lastAddedPath)
	} else if (FileArray.Length > 0) {
		SetStatus("No files added (duplicate or unsupported)", Symbol.Info)
	}
}

; Remove selected file from list
RemoveCustomFile(*) {
	global lst_Files, dllFileListChanged

	selectedIndices := lst_Files.Value ; Returns an array of indices in v2 multi-select ListBox
	if (selectedIndices.Length = 0) {
		MsgBox("Please select file(s) to remove.", "Info", "Icon!")
		return
	}

	; Delete from bottom to top to prevent index shifting
	loop selectedIndices.Length {
		idx := selectedIndices[selectedIndices.Length - A_Index + 1]
		lst_Files.Delete(idx)
	}

	dllFileListChanged := true
	SetStatus("File(s) removed", Symbol.Remove)
}

; Clear all files from list
ClearFileList(*) {
	global lst_Files, lv_Icons, CurrentDllPath, dllFileListChanged

	if (SendMessage(0x018B, 0, 0, lst_Files.Hwnd) == 0) ; LB_GETCOUNT
		return

	if (MsgBox("Are you sure you want to clear the file list?", "Clear List", "YesNo Icon?") = "No")
		return

	lst_Files.Delete()
	lv_Icons.Delete()
	CurrentDllPath := ""
	dllFileListChanged := true
	SetStatus("List cleared", Symbol.Trash)
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
	AddFilesToList(FileArray)
}

; Handle GUI resizing
GuiSize(GuiObj, MinMax, Width, Height) {
	if (MinMax = -1)
		return

	global mGui, lst_Files, lv_Icons, CurrentViewMode, pic_Preview, lbl_IconNo, edt_IconCode,
		lbl_PreviewSize,
		btn_CopyImage, btn_SaveImage, btn_CopyCode, btn_Test, ddl_ViewMode, lv_Favorites, btn_FavAdd, btn_FavRemove,
		lbl_Favorites, btn_FavClear, ddl_ExportSize, ExportIconSize, LastGuiW, LastGuiH, sb_Status

	LastGuiW := Width
	LastGuiH := Height

	GuiRedraw(GuiObj, 0) ; Disable redraw

	; Left panel
	lst_Files.Move(, , , Height - 110)

	; Right Panel
	rightPanelWidth := 210
	if (ExportIconSize + 20 > rightPanelWidth)
		rightPanelWidth := ExportIconSize + 20
	rightPanelX := Width - rightPanelWidth - 10

	sizeComboW := 70
	lblW := 100 ; Approximate width for the label text

	lbl_PreviewSize.Move(rightPanelX, 10, lblW)
	ddl_ExportSize.Move(rightPanelX + lblW + 5, 8, sizeComboW)

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

	favButtonsY := favoritesLabelY + 25
	btnWidth := 65
	btn_FavAdd.Move(rightPanelX, favButtonsY, btnWidth)
	btn_FavRemove.Move(rightPanelX + btnWidth + 5, favButtonsY, btnWidth)
	btn_FavClear.Move(rightPanelX + (btnWidth + 5) * 2, favButtonsY, btnWidth)

	favoritesListY := favButtonsY + 35
	favListHeight := Height - favoritesListY - 40
	if (favListHeight < 50)
		favListHeight := 50
	lv_Favorites.Move(rightPanelX, favoritesListY, rightPanelWidth, favListHeight)
	lv_Favorites.ModifyCol(3, rightPanelWidth - 38 - 50 - 10)

	; Middle Panel
	middlePanelStartX := 270 ; Starting X coordinate of IconListView
	middlePanelGapToRight := 10 ; Gap between the middle and right panels
	newListViewWidth := rightPanelX - middlePanelStartX - middlePanelGapToRight
	lv_Icons.Move(middlePanelStartX, 35, newListViewWidth, Height - 75)

	isReportView := CurrentViewMode < 2
	if (isReportView)
		lv_Icons.ModifyCol(1, newListViewWidth - 20)

	ddl_ViewMode.Move(middlePanelStartX + 65, 5)

	; Bottom status bar
	sb_Status.Move(, Height - 30, Width - 20)

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

; Sets the status bar text in its various panels.
; Panel 1: Main message/status
; Panel 2: Icon count info (if applicable)
; Panel 3: Current file path (if applicable)
; message - The main status message
; icon - Optional symbol from Symbol object
; iconInfo - Optional icon count string for Panel 2
; pathInfo - Optional file path for Panel 3
SetStatus(message, icon := "", iconInfo := "", pathInfo := "") {
	global sb_Status, Symbol, CurrentDllPath
	if (icon = "" && IsSet(Symbol))
		icon := Symbol.Info

	prefix := (icon != "") ? icon " " : ""
	sb_Status.SetText(prefix . message, 1)

	if (iconInfo != "")
		sb_Status.SetText(iconInfo, 2)

	if (pathInfo != "") {
		sb_Status.SetText(pathInfo, 3)
	} else if (CurrentDllPath != "") {
		sb_Status.SetText(CurrentDllPath, 3)
	}
}

; Show a tooltip for a short duration
ShowTempTooltip(message, icon := "", duration := 1500) {
	global Symbol
	if (icon = "" && IsSet(Symbol))
		icon := Symbol.Info
	prefix := (icon != "") ? icon " " : (Type(icon) = "String" ? icon " " : "")
	ToolTip(prefix . message, , , 1)
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