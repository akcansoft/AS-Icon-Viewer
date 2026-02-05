; AS Icon Viewer
; v1.01
; 05/02/2026

; Mesut Akcan
; -----------
; mesutakcan.blogspot.com
; github.com/akcansoft
; youtube.com/mesutakcan

#Requires AutoHotkey v2
#SingleInstance Force
#NoTrayIcon
#Include SaveFileDialog.ahk

; Global variables
A_ScriptName := "AS Icon Viewer v1.01"
;try TraySetIcon(A_WinDir "\System32\imageres.dll", 338)
try TraySetIcon(A_ScriptDir "\app_icon.ico")

global CurrentDllPath := ""
global CurrentViewMode := 3 ; 0:SmallReport, 1:LargeReport, 2:SmallIcon, 3:LargeIcon
global AllIcons := [] ; Array of all loaded icons
global IL_Small := 0 ; Small ImageList
global IL_Large := 0 ; Large ImageList
global SettingsFile := A_ScriptDir "\saved_files.txt"
global FavoritesFile := A_ScriptDir "\favorites.txt"
global lv_Favorites := 0
global IL_Favorites := 0
global FavoriteIcons := []
global dllFileListChanged := false
global FavoritesListChanged := false

global Txt := {
	Add: "➕ Add",
	Remove: "➖ Remove",
	Clear: "🧹 Clear",
	Save: "💾 Save",
	Test: "🧪 Test icon",
	Copy: "📋 Copy",
	CopyCode: "📋 Copy Code",
	AddFile: "➕ &Add File...",
	RemoveFile: "➖ &Remove File",
	ClearList: "🧹 &Clear List",
	Refresh: "&Refresh",
	SwitchView: "&Switch View",
	CopyImage: "📋 &Copy Image",
	SaveImage: "💾 &Save Image...",
	AddToFavorites: "➕ &Add to Favorites",
	RemoveFromFavorites: "➖ &Remove from Favorites"
}

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
mnu_Help.Add("&About...", (*) => ShowAbout())
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
btn_AddFile := mGui.AddButton("x10 y35 w80 h30", Txt.Add)
btn_AddFile.OnEvent("Click", AddCustomFile)
btn_RemoveFile := mGui.AddButton("x95 y35 w80 h30", Txt.Remove)
btn_RemoveFile.OnEvent("Click", RemoveCustomFile)
btn_ClearList := mGui.AddButton("x180 y35 w80 h30", Txt.Clear)
btn_ClearList.OnEvent("Click", ClearFileList)
lst_Files := mGui.AddListBox("x10 y70 w250 h465 +VScroll +HScroll", [])

; ========== MIDDLE PANEL (x270-740) ==========
mGui.AddText("x270 y10 w300", "🎨 Icons:")

; View selection - Radio buttons
mGui.SetFont("s11", "Segoe MDL2 Assets")
btn_Switch := mGui.AddButton("x600 y5 w140 h30", Chr(0xE8EB) " Switch View")
btn_Switch.OnEvent("Click", SwitchView)
mGui.SetFont("s9", "Segoe UI")

; Icons ListView
lv_Icons := mGui.AddListView("x270 y35 w400 h470 Grid", ["Icon and Number"])
lv_Icons.ModifyCol(1, 380)

; Right-click menu for Icons (Item selected)
mnu_IconContext := Menu()
mnu_IconContext.Add(Txt.CopyImage, (*) => CopyIconToClipboard())
mnu_IconContext.Add(Txt.SaveImage, (*) => SaveIconToFile())
mnu_IconContext.Add() ; Separator
mnu_IconContext.Add(Txt.CopyCode, (*) => CopyCurrentCode())
mnu_IconContext.Add(Txt.Test, (*) => TestIcon())
mnu_IconContext.Add() ; Separator
mnu_IconContext.Add("Add to Favorites", (*) => AddToFavorites())
mnu_IconContext.Add(Txt.Refresh, (*) => LoadIcons())
mnu_IconContext.Add(Txt.SwitchView, (*) => SwitchView())

; Right-click menu for Icons (Empty space)
mnu_IconEmptyContext := Menu()
mnu_IconEmptyContext.Add(Txt.Refresh, (*) => LoadIcons())
mnu_IconEmptyContext.Add(Txt.SwitchView, (*) => SwitchView())

; Right-click menu for File List
mnu_FileContext := Menu()
mnu_FileContext.Add(Txt.AddFile, (*) => AddCustomFile())
mnu_FileContext.Add(Txt.RemoveFile, (*) => RemoveCustomFile())
mnu_FileContext.Add(Txt.ClearList, (*) => ClearFileList())

lv_Icons.OnEvent("ContextMenu", ShowContextMenu)
lv_Icons.OnEvent("DoubleClick", CopyCurrentCode)
lv_Icons.OnEvent("ItemSelect", ShowPreview)

lst_Files.OnEvent("ContextMenu", ShowFileContextMenu)

; ========== RIGHT PANEL (x680-900) ==========
lbl_Preview := mGui.AddText("x680 y10 w210", "🔍 Preview:")
lbl_IconNo := mGui.AddText("x680 y35 w128 Center", "") ; Icon number text above picture
pic_Preview := mGui.AddPicture("x680 y55 w128 h128 Border")

btn_CopyImage := mGui.AddButton("x680 y190 w100 h30", Txt.Copy)
btn_CopyImage.OnEvent("Click", (*) => CopyIconToClipboard())

btn_SaveImage := mGui.AddButton("x+5 y190 w100 h30", Txt.Save)
btn_SaveImage.OnEvent("Click", (*) => SaveIconToFile())

edt_IconCode := mGui.AddEdit("x680 y230 w210 h50 -Wrap -VScroll", "")

btn_CopyCode := mGui.AddButton("x680 y285 w95 h30", Txt.CopyCode)
btn_CopyCode.OnEvent("Click", (*) => CopyCurrentCode())

btn_Test := mGui.AddButton("x+5 y285 w95 h30", Txt.Test)
btn_Test.OnEvent("Click", (*) => TestIcon())

lbl_Favorites := mGui.AddText("x680 y330 w210", "⭐ Favorites:")
lv_Favorites := mGui.AddListView("x680 y350 w210 h150 Grid", ["Icon", "Num", "File"])
lv_Favorites.ModifyCol(1, 38), lv_Favorites.ModifyCol(2, "50 Integer"), lv_Favorites.ModifyCol(3, 100)
lv_Favorites.OnEvent("ItemSelect", ShowFavoritePreview)

btn_FavAdd := mGui.AddButton("x680 y505 w65 h30", Txt.Add)
btn_FavAdd.OnEvent("Click", AddToFavorites)
btn_FavRemove := mGui.AddButton("x+5 y505 w65 h30", Txt.Remove)
btn_FavRemove.OnEvent("Click", RemoveFromFavorites)
btn_FavClear := mGui.AddButton("x+5 y505 w65 h30", Txt.Clear)
btn_FavClear.OnEvent("Click", ClearFavorites)

; ========== BOTTOM PANEL ==========
lbl_Status := mGui.AddText("x10 y545 w880", "💡 Select a file...")

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
InitializeDllList()
LoadFavorites()

; Event handlers
lst_Files.OnEvent("Change", LoadIcons)

; Show GUI
mGui.OnEvent("Close", (*) => CloseApplication())
mGui.OnEvent("Size", GuiSize)
mGui.OnEvent("DropFiles", OnDropFiles)
mGui.Show()
return

; ========== FUNCTIONS ==========

; Initialize DLL list with default files
InitializeDllList() {
	global lst_Files, DefaultDllFiles, SettingsFile, dllFileListChanged
	; Load saved files from settings
	if FileExist(SettingsFile) {
		try {
			savedFiles := FileRead(SettingsFile)
			Loop Parse, savedFiles, "`n", "`r" {
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
	global lv_Icons, lst_Files, lbl_Status, CurrentDllPath, AllIcons, CurrentViewMode
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
	AllIcons := []

	; Destroy previous ImageLists to free up memory
	try IL_Destroy(IL_Small)
	try IL_Destroy(IL_Large)

	; Create new ImageLists (Initial capacity: 100, Grow: 5)
	IL_Small := IL_Create(100, 5, false)
	IL_Large := IL_Create(100, 5, true)

	lbl_Status.Text := "⏳ Loading icons: " . CurrentDllPath . " ..."

	; --- DYNAMIC ICON COUNTING ---
	; Call Windows API to get the total number of icons in the file
	; Passing 0 to pHicon and pIconId returns only the icon count
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

	; Loop through the total number of icons found
	Loop iconCountTotal {
		iconIndex := A_Index

		; Add icon to small and large ImageLists
		hSmall := IL_Add(IL_Small, CurrentDllPath, iconIndex)
		hLarge := IL_Add(IL_Large, CurrentDllPath, iconIndex)

		; Skip if icon cannot be loaded (some indices might be empty or corrupt)
		if (hSmall = 0 && hLarge = 0)
			continue

		iconLoadedCount++
		AllIcons.Push({ num: iconIndex })

		; Update status bar progress every 20 icons for better UI responsiveness
		if (Mod(iconLoadedCount, 20) = 0) {
			lbl_Status.Text := "⏳ Loaded: " . iconLoadedCount . " / " . iconCountTotal . " icons..."
		}
	}

	; Default to "Large Icon" view (CurrentViewMode 3)
	CurrentViewMode := 3
	lv_Icons.Opt("+Icon")
	lv_Icons.SetImageList(IL_Large, 0)

	; Batch add rows to the ListView
	Loop iconLoadedCount {
		lv_Icons.Add("Icon" . A_Index, "  #" . AllIcons[A_Index].num)
	}

	; Final status update
	lbl_Status.Text := "✅ " . iconLoadedCount . " icons loaded | File: " . CurrentDllPath
}

; Show preview of selected icon
ShowPreview(*) {
	global lv_Icons, CurrentDllPath, pic_Preview, lbl_IconNo, edt_IconCode

	selectedRow := lv_Icons.GetNext(0, "Focused")
	if (selectedRow = 0) {
		pic_Preview.Value := ""
		lbl_IconNo.Text := ""
		edt_IconCode.Value := ""
		return
	}

	iconNum := GetIconNumberFromRow(selectedRow)

	; Show large icon (128x128)
	try {
		pic_Preview.Value := "*Icon" iconNum " *w128 *h128 " CurrentDllPath
	} catch {
		pic_Preview.Value := ""
	}

	; Show icon number overlay on picture
	lbl_IconNo.Text := "Icon #" iconNum

	; Update code preview based on selected language
	UpdateCodePreview()
}

; Update code preview based on selected language
UpdateCodePreview(*) {
	global lv_Icons, CurrentDllPath, edt_IconCode

	selectedRow := lv_Icons.GetNext(0, "Focused")
	if (selectedRow = 0 || CurrentDllPath = "") {
		edt_IconCode.Value := ""
		return
	}

	iconNum := GetIconNumberFromRow(selectedRow)
	SplitPath(CurrentDllPath, &fileName)

	edt_IconCode.Value := 'TraySetIcon("' fileName '", ' iconNum ')'
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

	text := lv_Icons.GetText(rowNumber, 1)
	text := StrReplace(text, " ", "")
	text := StrReplace(text, "#", "")
	return Integer(text)
}

; Copy current icon code to clipboard
CopyCurrentCode(*) {
	global edt_IconCode, lbl_Status

	code := edt_IconCode.Value
	if (code = "") {
		ToolTip("⚠️ Select an icon first!", , , 1)
		SetTimer(() => ToolTip(, , , 1), -1500)
		return
	}

	A_Clipboard := code
	ToolTip("📋 Code copied!", , , 1)
	SetTimer(() => ToolTip(, , , 1), -1500)
	lbl_Status.Text := "📋 Copied: " code
}

; Copy icon image to clipboard
CopyIconToClipboard(*) {
	global CurrentDllPath, lv_Icons, lbl_Status

	; Get selected row
	selectedRow := lv_Icons.GetNext(0, "Focused")
	if (selectedRow = 0) {
		ToolTip("⚠️ Select an icon first!", , , 1)
		SetTimer(() => ToolTip(, , , 1), -1500)
		return
	}

	iconNum := GetIconNumberFromRow(selectedRow)

	try {
		; Load icon as HICON
		hIcon := LoadPicture(CurrentDllPath, "Icon" iconNum " w128 h128", &imageType)

		if (!hIcon) {
			ToolTip("❌ Failed to load icon!", , , 1)
			SetTimer(() => ToolTip(, , , 1), -1500)
			return
		}

		; Create a device context
		hdc := DllCall("GetDC", "Ptr", 0, "Ptr")
		hdcMem := DllCall("CreateCompatibleDC", "Ptr", hdc, "Ptr")

		; Create a bitmap
		hBitmap := DllCall("CreateCompatibleBitmap", "Ptr", hdc, "Int", 128, "Int", 128, "Ptr")
		hOldBitmap := DllCall("SelectObject", "Ptr", hdcMem, "Ptr", hBitmap, "Ptr")

		; Fill background with white
		hBrush := DllCall("CreateSolidBrush", "UInt", 0xFFFFFF, "Ptr")
		DllCall("FillRect", "Ptr", hdcMem, "Ptr", Buffer(16, 0), "Ptr", hBrush)
		DllCall("DeleteObject", "Ptr", hBrush)

		; Draw icon on bitmap
		DllCall("DrawIconEx", "Ptr", hdcMem, "Int", 0, "Int", 0, "Ptr", hIcon, "Int", 128, "Int", 128, "UInt", 0, "Ptr", 0, "UInt", 0x0003) ; DI_NORMAL

		; Cleanup
		DllCall("SelectObject", "Ptr", hdcMem, "Ptr", hOldBitmap)
		DllCall("DeleteDC", "Ptr", hdcMem)
		DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdc)
		DllCall("DestroyIcon", "Ptr", hIcon)

		; Open clipboard and set bitmap
		if (DllCall("OpenClipboard", "Ptr", 0)) {
			DllCall("EmptyClipboard")
			DllCall("SetClipboardData", "UInt", 2, "Ptr", hBitmap) ; CF_BITMAP = 2
			DllCall("CloseClipboard")

			ToolTip("✅ Icon copied to clipboard!", , , 1)
			SetTimer(() => ToolTip(, , , 1), -1500)
			lbl_Status.Text := "📋 Icon image copied to clipboard"
		} else {
			DllCall("DeleteObject", "Ptr", hBitmap)
			ToolTip("❌ Failed to open clipboard!", , , 1)
			SetTimer(() => ToolTip(, , , 1), -1500)
		}

	} catch as err {
		ToolTip("❌ Error: " err.Message, , , 1)
		SetTimer(() => ToolTip(, , , 1), -2000)
	}
}

; Save current icon to file (PNG)
SaveIconToFile(*) {
	global CurrentDllPath, lv_Icons, lbl_Status

	selectedRow := lv_Icons.GetNext(0, "Focused")
	if (selectedRow = 0) {
		ToolTip("⚠️ Select an icon first!", , , 1)
		SetTimer(() => ToolTip(, , , 1), -1500)
		return
	}

	iconNum := GetIconNumberFromRow(selectedRow)

	SplitPath(CurrentDllPath, &fileName)
	safeFileName := StrReplace(fileName, ".", "_")
	defaultSaveName := safeFileName "_Icon_" iconNum ".ico"
	saved := SaveFile([mGui.Hwnd, "Save Icon"]
		, defaultSaveName
		, { ICO: "*.ico`n", PNG: "*.png", BMP: "*.bmp", JPEG: "*.jpg" }
		, ""
		, 2)

	if !saved
		return
	savePath := saved.FileFullPath

	try {
		hIcon := LoadPicture(CurrentDllPath, "Icon" iconNum " w128 h128", &imageType)
		if (!hIcon)
			throw Error("Failed to load icon")

		SaveHICONToFile(hIcon, savePath)
		DllCall("DestroyIcon", "Ptr", hIcon)

		lbl_Status.Text := "✅ Saved: " savePath
		ToolTip("✅ Icon saved!", , , 1)
		SetTimer(() => ToolTip(, , , 1), -1500)
	} catch as err {
		MsgBox("Error saving image: " err.Message, "Error", "Icon!")
	}
}

; Helper to save HICON to various formats using GDI+
SaveHICONToFile(hIcon, filePath) {
	SplitPath(filePath, , , &ext)
	ext := StrLower(ext)

	; CLSIDs for image formats
	clsids := Map(
		"png", "{557CF406-1A04-11D3-9A73-0000F81EF32E}",
		"bmp", "{557CF400-1A04-11D3-9A73-0000F81EF32E}",
		"jpg", "{557CF401-1A04-11D3-9A73-0000F81EF32E}",
		"jpeg", "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
	)

	hModule := DllCall("LoadLibrary", "Str", "gdiplus", "Ptr")
	si := Buffer(24, 0), NumPut("UInt", 1, si)
	DllCall("gdiplus\GdiplusStartup", "Ptr*", &pToken := 0, "Ptr", si, "Ptr", 0)
	DllCall("gdiplus\GdipCreateBitmapFromHICON", "Ptr", hIcon, "Ptr*", &pBitmap := 0)

	if (ext = "ico") {
		; For ICO, we save as PNG first (to preserve transparency) and wrap it in an ICO container
		tempFile := A_Temp "\temp_icon_" A_TickCount ".png"
		DllCall("ole32\CLSIDFromString", "Str", clsids["png"], "Ptr", Encoder := Buffer(16))
		DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "WStr", tempFile, "Ptr", Encoder, "Ptr", 0)

		try {
			pngData := FileRead(tempFile, "RAW")
			f := FileOpen(filePath, "w")
			; ICONDIR
			f.WriteUShort(0), f.WriteUShort(1), f.WriteUShort(1)
			; ICONDIRENTRY
			f.WriteUChar(128) ; Width
			f.WriteUChar(128) ; Height
			f.WriteUChar(0), f.WriteUChar(0)
			f.WriteUShort(1), f.WriteUShort(32)
			f.WriteUInt(pngData.Size)
			f.WriteUInt(22) ; Offset (6+16)
			f.RawWrite(pngData)
			f.Close()
			FileDelete(tempFile)
		}
	} else if clsids.Has(ext) {
		DllCall("ole32\CLSIDFromString", "Str", clsids[ext], "Ptr", Encoder := Buffer(16))
		DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "WStr", filePath, "Ptr", Encoder, "Ptr", 0)
	}

	DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap), DllCall("gdiplus\GdiplusShutdown", "Ptr", pToken), DllCall("FreeLibrary", "Ptr", hModule)
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
	global lv_Icons, CurrentDllPath, mGui

	selectedRow := lv_Icons.GetNext(0, "Focused")
	if (selectedRow = 0) {
		MsgBox("Please select an icon from the list to test.", "Warning", "Icon!")
		return
	}

	iconNum := GetIconNumberFromRow(selectedRow)

	try {
		; Change tray icon
		TraySetIcon(CurrentDllPath, iconNum)

		; Change window icon (WM_SETICON)
		if (hIconSmall := LoadPicture(CurrentDllPath, "Icon" . iconNum . " w16 h16", &isIcon))
			SendMessage(0x80, 0, hIconSmall, mGui.Hwnd) ; ICON_SMALL
		if (hIconBig := LoadPicture(CurrentDllPath, "Icon" . iconNum . " w32 h32", &isIcon))
			SendMessage(0x80, 1, hIconBig, mGui.Hwnd) ; ICON_BIG

		ToolTip("✅ Application icon updated")
		SetTimer(() => ToolTip(), -2000)
	} catch as err {
		MsgBox("Error changing icon: " err.Message, "Error", "Icon!")
	}
}

; Add current icon to favorites
AddToFavorites(*) {
	global lv_Icons, CurrentDllPath, lbl_Status, lv_Favorites, IL_Favorites, FavoriteIcons, FavoritesListChanged

	selectedRow := lv_Icons.GetNext(0, "Focused")
	if (selectedRow = 0) {
		ToolTip("⚠️ Select an icon first!", , , 1)
		SetTimer(() => ToolTip(, , , 1), -1500)
		return
	}

	iconNum := GetIconNumberFromRow(selectedRow)
	filePath := CurrentDllPath

	for fav in FavoriteIcons {
		if (fav.path = filePath && fav.num = iconNum) {
			ToolTip("⚠️ Icon already in favorites!", , , 1)
			SetTimer(() => ToolTip(, , , 1), -1500)
			return
		}
	}

	FavoriteIcons.Push({ path: filePath, num: iconNum })
	FavoritesListChanged := true

	iconIndex := IL_Add(IL_Favorites, filePath, iconNum)
	SplitPath(filePath, &fileName)
	lv_Favorites.Add("Icon" . iconIndex, "", iconNum, fileName)

	lbl_Status.Text := "⭐ Added to favorites: " fileName " #" iconNum
	ToolTip("⭐ Added to favorites!", , , 1)
	SetTimer(() => ToolTip(, , , 1), -1500)
}

RemoveFromFavorites(*) {
	global lv_Favorites, FavoriteIcons, FavoritesListChanged

	selected := lv_Favorites.GetNext(0, "Focused")
	if (selected = 0) {
		ToolTip("⚠️ Select a favorite to remove!", , , 1)
		SetTimer(() => ToolTip(, , , 1), -1500)
		return
	}

	FavoriteIcons.RemoveAt(selected)
	lv_Favorites.Delete(selected)
	FavoritesListChanged := true

	ToolTip("➖ Favorite removed.", , , 1)
	SetTimer(() => ToolTip(, , , 1), -1500)
}

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
	ToolTip("🗑️ Favorites cleared", , , 1)
	SetTimer(() => ToolTip(, , , 1), -1500)
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
		Loop Read, FavoritesFile {
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
	global lv_Favorites, FavoriteIcons, pic_Preview, lbl_IconNo, edt_IconCode

	selected := lv_Favorites.GetNext(0, "Focused")
	if (selected = 0)
		return

	fav := FavoriteIcons[selected]

	pic_Preview.Value := "*Icon" fav.num " *w128 *h128 " fav.path
	lbl_IconNo.Text := "Icon #" fav.num

	SplitPath(fav.path, &fileName)
	edt_IconCode.Value := 'TraySetIcon("' fileName '", ' fav.num ')'
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
	itemCount := SendMessage(0x018B, 0, 0, lst_Files.Hwnd) ; LB_GETCOUNT
	Loop itemCount {
		textBuf := Buffer(1024)
		SendMessage(0x0189, A_Index - 1, textBuf, lst_Files.Hwnd) ; LB_GETTEXT
		if (StrGet(textBuf) = filePathToAdd) {
			MsgBox("This file is already in the list!", "Info", "Iconi")
			return
		}
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
	global lst_Files, lv_Icons, CurrentDllPath, lbl_Status, AllIcons, dllFileListChanged

	if (SendMessage(0x018B, 0, 0, lst_Files.Hwnd) == 0) ; LB_GETCOUNT
		return

	if (MsgBox("Are you sure you want to clear the file list?", "Clear List", "YesNo Icon?") = "No")
		return

	lst_Files.Delete()
	lv_Icons.Delete()
	AllIcons := []
	CurrentDllPath := ""
	dllFileListChanged := true
	lbl_Status.Text := "🗑️ List cleared"
}

; Save file list and exit
CloseApplication(*) {
	SaveFileList()
	SaveFavorites()
	ExitApp()
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
		alreadyExists := false
		itemCount := SendMessage(0x018B, 0, 0, lst_Files.Hwnd) ; LB_GETCOUNT
		Loop itemCount {
			textBuf := Buffer(1024)
			SendMessage(0x0189, A_Index - 1, textBuf, lst_Files.Hwnd) ; LB_GETTEXT
			if (StrGet(textBuf) = filePathToAdd) {
				alreadyExists := true
				break
			}
		}

		if !alreadyExists {
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

	global mGui, lst_Files, lv_Icons, lbl_Status, CurrentViewMode, pic_Preview, lbl_IconNo, edt_IconCode, lbl_Preview, btn_CopyImage, btn_SaveImage, btn_CopyCode, btn_Test, btn_Switch, lv_Favorites, btn_FavAdd, btn_FavRemove, lbl_Favorites, btn_FavClear

	GuiRedraw(GuiObj, 0) ; Disable redraw to prevent flickering

	; Left panel
	lst_Files.Move(, , , Height - 110)

	; Right Panel
	rightPanelWidth := 210
	rightPanelX := Width - rightPanelWidth - 10
	lbl_Preview.Move(rightPanelX), lbl_IconNo.Move(rightPanelX), pic_Preview.Move(rightPanelX)
	btn_CopyImage.Move(rightPanelX), btn_SaveImage.Move(rightPanelX + 105)
	edt_IconCode.Move(rightPanelX, 230, rightPanelWidth)
	btn_CopyCode.Move(rightPanelX), btn_Test.Move(rightPanelX + 105)
	lbl_Favorites.Move(rightPanelX, 330)

	favButtonsY := Height - 70
	favListHeight := favButtonsY - 350 - 5
	lv_Favorites.Move(rightPanelX, 350, rightPanelWidth, favListHeight)
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