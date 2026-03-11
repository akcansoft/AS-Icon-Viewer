;@Ahk2Exe-SetName AS Icon Viewer
;@Ahk2Exe-SetDescription A professional utility to browse`, preview`, and export icons from DLL`, EXE`, and ICO files
;@Ahk2Exe-SetFileVersion 1.4
;@Ahk2Exe-SetCompanyName AkcanSoft
;@Ahk2Exe-SetCopyright ©2026 Mesut Akcan
;@Ahk2Exe-SetMainIcon app_icon.ico

; AS Icon Viewer
; Mesut Akcan
; 11/03/2026
; -------------------------------------------
; A high-performance utility to browse, preview, and export
; icons from DLL, EXE, and ICO files.
;
; Author: Mesut Akcan
; Home: mesutakcan.blogspot.com
; Source: github.com/akcansoft/AS-Icon-Viewer
; YouTube: youtube.com/mesutakcan

#Requires AutoHotkey v2
#SingleInstance Force

#Include Gdip.ahk ; GDI+
#Include SaveFileDialog.ahk ; Custom Save File Dialog

; ========== TRAY ICON ==========
; Set custom icon for the tray menu
if (!A_IsCompiled)
	try TraySetIcon(A_ScriptDir "\app_icon.ico")

; Global variables
global App := {
	Name: "AS Icon Viewer",
	Version: "1.4",
	DllListFile: A_ScriptDir "\saved_files.txt",
	FavoritesFile: A_ScriptDir "\favorites.txt",
	WebUrl: "https://mesutakcan.blogspot.com",
	GitHubUrl: "https://github.com/akcansoft/AS-Icon-Viewer"
}

global IL := {
	Small: 0,
	Large: 0,
	Favorites: 0
}

global State := {
	CurrentDllPath: "",       ; Loaded file path
	CurrentViewMode: 3,       ; 0:SmallReport, 1:LargeReport, 2:SmallIcon, 3:LargeIcon
	dllFileListChanged: false, ; Track file list persistence
	FavoritesListChanged: false, ; Track favorites persistence
	ExportIconSize: 128,      ; Current preview/export dimension
	LastGuiW: 0,              ; UI resize state
	LastGuiH: 0,
	Interpolation: 1          ; HQ scaling vs Nearest-Neighbor
}
global PreviewIcon := { Path: "", Num: 0 } ; Active preview metadata
global g_hPreviewBitmap := 0               ; Cache for STM_SETIMAGE rendering
global g_hWndIconSmall := 0                ; Window icon handle (small)
global g_hWndIconBig := 0                  ; Window icon handle (big)
global FilePathMap := Map()                ; Hash map for O(1) deduplication

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
	Web: "🌐",
	Exit: "✖️"
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
	Exit: Symbol.Exit " E&xit",
	Refresh: "&Refresh",
	CopyImage: Symbol.Copy " Co&py Image",
	SaveImage: Symbol.Save " &Save Image ",
	AddToFavorites: Symbol.Add " &Add to Favorites",
	RemoveFromFavorites: Symbol.Remove " &Remove from Favorites",
	About: Symbol.About " &About",
	Website: Symbol.Web " &My Blog",
	GitHub: Symbol.Web " &GitHub Repository",
	GoToSource: "🔍 &Go to Source",
	Properties: "📋 &Properties",
	View0: "&Small Report",
	View1: "&Large Report",
	View2: "S&mall Icon",
	View3: "L&arge Icon",
	Interpolation: "&Interpolation on"
}

; --- DPI-Aware Metrics ---
global Metric := {
	SmallIcon: {
		W: DllCall("User32.dll\GetSystemMetrics", "Int", 49, "Int"), ; SM_CXSMICON
		H: DllCall("User32.dll\GetSystemMetrics", "Int", 50, "Int")  ; SM_CYSMICON
	},
	LargeIcon: {
		W: DllCall("User32.dll\GetSystemMetrics", "Int", 11, "Int"), ; SM_CXICON
		H: DllCall("User32.dll\GetSystemMetrics", "Int", 12, "Int")  ; SM_CYICON
	}
}

global Layout := {
	Margin: 10,
	Left: {
		X: 10,
		W: 250,
		BtnW: 80,
		BtnH: 30,
		ListY: 70,
		BottomOffset: 110
	},
	Middle: {
		X: 270,
		Gap: 10,
		DdlXOffset: 65
	},
	Right: {
		W: 210,
		BtnW: 100,
		BtnH: 30,
		TestBtnXOffset: 105,
		FavBtnW: 65,
		SizeW: 70,
		LabelW: 100,
		FavIconColW: 38,
		FavNumColW: 50
	}
}

; ========== TRAY MENU ==========
A_TrayMenu.Delete()
A_TrayMenu.Add(Txt.About, (*) => ShowAbout())
A_TrayMenu.Add(Txt.Website, (*) => Run(App.WebUrl))
A_TrayMenu.Add(Txt.GitHub, (*) => Run(App.GitHubUrl))
A_TrayMenu.Add()
A_TrayMenu.Add "Open " App.Name, (*) => g.Show()
A_TrayMenu.Add()
A_TrayMenu.Add(Txt.Exit, (*) => CloseApplication())
A_TrayMenu.Default := "Open " App.Name

; Create main GUI
g := Gui("+Resize +MinSize900x600", App.Name)
g.SetFont("s9", "Segoe UI")

; ========== MENU BAR ==========
; File Menu
mnu_File := Menu()
mnu_File.Add(Txt.AddFile, (*) => AddCustomFile())
mnu_File.Add(Txt.RemoveFile, (*) => RemoveCustomFile())
mnu_File.Add(Txt.ClearList, (*) => ClearFileList())
mnu_File.Add()
mnu_File.Add(Txt.Exit, (*) => CloseApplication())

; Icon View Menu
mnu_View := Menu()
mnu_View.Add(Txt.View0, (*) => SetViewMode(0))
mnu_View.Add(Txt.View1, (*) => SetViewMode(1))
mnu_View.Add(Txt.View2, (*) => SetViewMode(2))
mnu_View.Add(Txt.View3, (*) => SetViewMode(3))
mnu_View.Add()
mnu_View.Add(Txt.Refresh, (*) => LoadIcons(true))

; Icon Menu
mnu_Icon := Menu()
mnu_Icon.Add(Txt.CopyImage, (*) => CopyIconToClipboard())
mnu_Icon.Add(Txt.SaveImage, (*) => SaveIconToFile())
mnu_Icon.Add(Txt.CopyCode, (*) => CopyCurrentCode())
mnu_Icon.Add(Txt.Test, (*) => TestIcon())
mnu_Icon.Add()
mnu_Icon.Add(Txt.Interpolation, (*) => ToggleInterpolation())
if (State.Interpolation)
	mnu_Icon.Check(Txt.Interpolation) ; Default: On

; Favorites Menu
mnu_Favorites := Menu()
mnu_Favorites.Add(Txt.AddToFavorites, (*) => AddToFavorites())
mnu_Favorites.Add(Txt.RemoveFromFavorites, (*) => RemoveFromFavorites())
mnu_Favorites.Add()
mnu_Favorites.Add(Txt.GoToSource, (*) => GoToSource())
mnu_Favorites.Add(Txt.OpenFolder, (*) => OpenFileLocation())
mnu_Favorites.Add(Txt.Properties, (*) => ShowFileProperties())
mnu_Favorites.Add()
mnu_Favorites.Add(Txt.ClearFav, (*) => ClearFavorites())

; Help Menu
mnu_Help := Menu()
mnu_Help.Add(Txt.About, (*) => ShowAbout())
mnu_Help.Add()
mnu_Help.Add(Txt.Website, (*) => Run(App.WebUrl))
mnu_Help.Add(Txt.GitHub, (*) => Run(App.GitHubUrl))

; Main Menu Bar
mnu_Main := MenuBar()
mnu_Main.Add(Txt.File, mnu_File)
mnu_Main.Add(Txt.View, mnu_View)
mnu_Main.Add(Txt.Favorites, mnu_Favorites)
mnu_Main.Add(Txt.Icon, mnu_Icon)
mnu_Main.Add(Txt.Help, mnu_Help)
g.MenuBar := mnu_Main

; Right-click menu for Icons (Item selected)
mnu_IconContext := Menu()
mnu_IconContext.Add(Txt.CopyImage, (*) => CopyIconToClipboard())
mnu_IconContext.Add(Txt.SaveImage, (*) => SaveIconToFile())
mnu_IconContext.Add()
mnu_IconContext.Add(Txt.CopyCode, (*) => CopyCurrentCode())
mnu_IconContext.Add(Txt.Test, (*) => TestIcon())
mnu_IconContext.Add()
mnu_IconContext.Add(Txt.AddToFavorites, (*) => AddToFavorites())
mnu_IconContext.Add(Txt.Refresh, (*) => LoadIcons(true))

; Right-click menu for Icons (Empty space)
mnu_IconEmptyContext := Menu()
mnu_IconEmptyContext.Add(Txt.Refresh, (*) => LoadIcons(true))

; Right-click menu for File List (item selected)
mnu_FileContext_Item := Menu()
mnu_FileContext_Item.Add(Txt.OpenFolder, (*) => OpenFileLocation())
mnu_FileContext_Item.Add(Txt.RemoveFile, (*) => RemoveCustomFile())
mnu_FileContext_Item.Add()
mnu_FileContext_Item.Add(Txt.AddFile, (*) => AddCustomFile())
mnu_FileContext_Item.Add(Txt.ClearList, (*) => ClearFileList())

; Right-click menu for File List (empty space)
mnu_FileContext_Empty := Menu()
mnu_FileContext_Empty.Add(Txt.AddFile, (*) => AddCustomFile())
mnu_FileContext_Empty.Add(Txt.ClearList, (*) => ClearFileList())

; Right-click menu for Favorites List (item selected)
mnu_FavContext_Item := Menu()
mnu_FavContext_Item.Add(Txt.RemoveFav, (*) => RemoveFromFavorites())
mnu_FavContext_Item.Add()
mnu_FavContext_Item.Add(Txt.CopyImage, (*) => CopyIconToClipboard())
mnu_FavContext_Item.Add(Txt.SaveImage, (*) => SaveIconToFile())
mnu_FavContext_Item.Add(Txt.Test, (*) => TestIcon())
mnu_FavContext_Item.Add(Txt.CopyCode, (*) => CopyCurrentCode())
mnu_FavContext_Item.Add()
mnu_FavContext_Item.Add(Txt.GoToSource, (*) => GoToSource())
mnu_FavContext_Item.Add(Txt.OpenFolder, (*) => OpenFileLocation())
mnu_FavContext_Item.Add(Txt.Properties, (*) => ShowFileProperties())

; Right-click menu for Favorites List (empty space)
mnu_FavContext_Empty := Menu()
mnu_FavContext_Empty.Add(Txt.AddFav, (*) => AddToFavorites())
mnu_FavContext_Empty.Add(Txt.ClearFav, (*) => ClearFavorites())

; --- Sidebar (File List) ---
g.AddText("x" Layout.Left.X " y10 w" Layout.Left.W, Symbol.File " Icon Files:")
btn_AddFile := g.AddButton("x" Layout.Left.X " y35 w" Layout.Left.BtnW " h" Layout.Left.BtnH, Txt.Add)
btn_AddFile.OnEvent("Click", AddCustomFile)
btn_RemoveFile := g.AddButton("x" (Layout.Left.X + 85) " y35 w" Layout.Left.BtnW " h" Layout.Left.BtnH, Txt.Remove)
btn_RemoveFile.OnEvent("Click", RemoveCustomFile)
btn_ClearList := g.AddButton("x" (Layout.Left.X + 170) " y35 w" Layout.Left.BtnW " h" Layout.Left.BtnH, Txt.Clear)
btn_ClearList.OnEvent("Click", ClearFileList)
lst_Files := g.AddListBox("x" Layout.Left.X " y" Layout.Left.ListY " w" Layout.Left.W " h465 Multi +VScroll +HScroll",
	[])

; ========== MIDDLE PANEL ==========
g.AddText("x" Layout.Middle.X " y10", Symbol.Color " Icons:")
ddl_ViewMode := g.AddDropDownList("x" (Layout.Middle.X + Layout.Middle.DdlXOffset) " y5", [StrReplace(Txt.View0, "&"),
	StrReplace(Txt.View1,
		"&"), StrReplace(Txt
			.View2, "&"), StrReplace(Txt.View3, "&")])
ddl_ViewMode.OnEvent("Change", OnViewChange)

; Icons ListView
lv_Icons := g.AddListView("x" Layout.Middle.X " y35 w400 h470 Grid -Multi", ["Icon and Number"])
lv_Icons.ModifyCol(1, 380)

; Icons ListView Event Handlers
lv_Icons.OnEvent("ContextMenu", ShowContextMenu)
lv_Icons.OnEvent("DoubleClick", CopyCurrentCode)
lv_Icons.OnEvent("ItemSelect", (*) => ShowPreview())
lv_Icons.OnEvent("Click", (*) => ShowPreview(false))

; File List Event Handlers
lst_Files.OnEvent("ContextMenu", ShowFileContextMenu)

; --- Preview & Actions ---
lbl_PreviewSize := g.AddText("x680 y10", Symbol.Search " Preview Size:")
ddl_ExportSize := g.AddDropDownList("x680 y8 w" Layout.Right.SizeW, ["16", "24", "32", "48", "64", "128", "256"])
ddl_ExportSize.Text := State.ExportIconSize
ddl_ExportSize.OnEvent("Change", OnExportSizeChange)
lbl_IconNo := g.AddText("x680 y35 w80", "")
pic_Preview := g.AddPicture("x680 y55 w128 h128 Border")
cb_Interpolation := g.AddCheckBox("x680 y190 w210", Txt.Interpolation)
cb_Interpolation.Value := State.Interpolation
cb_Interpolation.OnEvent("Click", OnInterpolationToggle)

btn_CopyImage := g.AddButton("x680 y215 w" Layout.Right.BtnW " h" Layout.Right.BtnH, Txt.Copy)
btn_CopyImage.OnEvent("Click", (*) => CopyIconToClipboard())

btn_SaveImage := g.AddButton("x+5 y215 w" Layout.Right.BtnW " h" Layout.Right.BtnH, Txt.Save)
btn_SaveImage.OnEvent("Click", (*) => SaveIconToFile())

edt_IconCode := g.AddEdit("x680 y230 w210 h50", "")

btn_CopyCode := g.AddButton("x680 y285 w95 h30", Txt.CopyCode)
btn_CopyCode.OnEvent("Click", (*) => CopyCurrentCode())

btn_Test := g.AddButton("x+5 y285 w95 h30", Txt.Test)
btn_Test.OnEvent("Click", (*) => TestIcon())

lbl_Favorites := g.AddText("x680 y335 w210", Symbol.Star " Favorites:")
btn_FavAdd := g.AddButton("x680 y360 w" Layout.Right.FavBtnW " h" Layout.Right.BtnH, Txt.AddFav)
btn_FavAdd.OnEvent("Click", AddToFavorites)
btn_FavRemove := g.AddButton("x+5 y360 w" Layout.Right.FavBtnW " h" Layout.Right.BtnH, Txt.RemoveFav)
btn_FavRemove.OnEvent("Click", RemoveFromFavorites)
btn_FavClear := g.AddButton("x+5 y360 w" Layout.Right.FavBtnW " h" Layout.Right.BtnH, Txt.ClearFav)
btn_FavClear.OnEvent("Click", ClearFavorites)

lv_Favorites := g.AddListView("x680 y395 w210 h150 Grid", ["Icon", "Num", "File", "FullPath"])
lv_Favorites.ModifyCol(2, "Integer"), lv_Favorites.ModifyCol(3, 100), lv_Favorites.ModifyCol(4, 0) ; Hidden path column
lv_Favorites.OnEvent("ItemSelect", (*) => ShowFavoritePreview())
lv_Favorites.OnEvent("Click", (*) => ShowFavoritePreview(false))
lv_Favorites.OnEvent("ContextMenu", ShowFavContextMenu)

; --- Status Bar ---
sb_Status := g.AddStatusBar()
sb_Status.SetParts(250, 150) ; Process | Count | Path
SetStatus("Select a file", Symbol.Info)

; Create ImageLists (initially using system metrics for high DPI sharpness)
IL.Small := CreateImageList(Metric.SmallIcon.W, Metric.SmallIcon.H)
IL.Large := CreateImageList(Metric.LargeIcon.W, Metric.LargeIcon.H)
IL.Favorites := CreateImageList(Metric.SmallIcon.W, Metric.SmallIcon.H)
lv_Favorites.SetImageList(IL.Favorites)

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
InitializeDllList() ; Load persistent file list or defaults
LoadFavorites()     ; Load favorites from file
Gdip.Startup()      ; Initialize GDI+
OnMessage(0x0100, OnLvKeyDown) ; Navigation handling
ApplyViewSettings() ; Set initial UI state

; Ensure GDI+ is shut down on any exit (crash, Windows shutdown, task kill)
OnExit((*) => Gdip.Shutdown())

; Event handlers
lst_Files.OnEvent("Change", (*) => LoadIcons())

; Show GUI
g.OnEvent("Close", (*) => CloseApplication())
g.OnEvent("Size", GuiSize)
g.OnEvent("DropFiles", OnDropFiles)
g.Show("w950 h670")
return

; ========== FUNCTIONS ==========
; Creates an ImageList with the given icon dimensions.
CreateImageList(w, h) {
	return DllCall("Comctl32.dll\ImageList_Create",
		"Int", w, "Int", h, "UInt", 0x21,
		"Int", 100, "Int", 5, "Ptr")
}

OnViewChange(*) {
	SetViewMode(ddl_ViewMode.Value - 1)
}

; Initialize DLL list with default files
InitializeDllList() {
	global DefaultDllFiles, State, FilePathMap
	SendMessage(0x000B, 0, 0, lst_Files.Hwnd) ; WM_SETREDRAW off
	; Load saved files from settings
	if FileExist(App.DllListFile) {
		try {
			savedFiles := FileRead(App.DllListFile)
			loop parse, savedFiles, "`n", "`r" {
				if (A_LoopField != "" && FileExist(A_LoopField)) {
					lst_Files.Add([A_LoopField])
					FilePathMap[A_LoopField] := true
				}
			}
		}
	}
	; If no saved files, load defaults
	if (ControlGetItems(lst_Files).Length == 0) {
		for dllName in DefaultDllFiles {
			dllPath := A_WinDir "\System32\" dllName
			if FileExist(dllPath) {
				lst_Files.Add([dllPath])
				FilePathMap[dllPath] := true
			}
		}
		State.dllFileListChanged := true
	}
	SendMessage(0x000B, 1, 0, lst_Files.Hwnd) ; WM_SETREDRAW on
	DllCall("InvalidateRect", "Ptr", lst_Files.Hwnd, "Ptr", 0, "Int", true)
}

; Load icons from the selected DLL/EXE/ICO file
LoadIcons(force := false, *) {
	global State, g

	; Check if a file is selected in the list
	selectedIndices := lst_Files.Value
	if (Type(selectedIndices) = "Array" ? selectedIndices.Length = 0 : selectedIndices = 0)
		return

	; Handle multi-select ListBox (Text returns an array)
	selectedFile := lst_Files.Text
	if (Type(selectedFile) = "Array")
		selectedFile := (selectedFile.Length > 0) ? selectedFile[1] : ""

	if (selectedFile = "")
		return
	if (!force && selectedFile = State.CurrentDllPath)
		return

	State.CurrentDllPath := selectedFile
	lv_Icons.Delete()
	try IL_Destroy(IL.Small)
	try IL_Destroy(IL.Large)
	IL.Small := CreateImageList(Metric.SmallIcon.W, Metric.SmallIcon.H)
	IL.Large := CreateImageList(Metric.LargeIcon.W, Metric.LargeIcon.H)
	SetStatus("Loading icons: " . State.CurrentDllPath, Symbol.Loading)
	iconCountTotal := DllCall("User32.dll\PrivateExtractIconsW",
		"Str", State.CurrentDllPath,
		"Int", 0, "Int", 0, "Int", 0,
		"Ptr", 0, "Ptr", 0,
		"UInt", 0, "UInt", 0, "UInt")
	if (iconCountTotal <= 0) {
		SetStatus("No icons found in: " . State.CurrentDllPath, Symbol.Error)
		return
	}

	; Batch extract all small and large icons to minimize disk I/O.
	hIconsSmall := Buffer(A_PtrSize * iconCountTotal, 0)
	hIconsLarge := Buffer(A_PtrSize * iconCountTotal, 0)
	DllCall("User32.dll\PrivateExtractIconsW",
		"Str", State.CurrentDllPath, "Int", 0,
		"Int", Metric.SmallIcon.W, "Int", Metric.SmallIcon.H,
		"Ptr", hIconsSmall, "Ptr", 0, "UInt", iconCountTotal, "UInt", 0)
	DllCall("User32.dll\PrivateExtractIconsW",
		"Str", State.CurrentDllPath, "Int", 0,
		"Int", Metric.LargeIcon.W, "Int", Metric.LargeIcon.H,
		"Ptr", hIconsLarge, "Ptr", 0, "UInt", iconCountTotal, "UInt", 0)

	iconLoadedCount := 0
	lv_Icons.Opt("-Redraw") ; Disable redraw during loading for performance
	loop iconCountTotal {
		hSmall := NumGet(hIconsSmall, (A_Index - 1) * A_PtrSize, "Ptr")
		hLarge := NumGet(hIconsLarge, (A_Index - 1) * A_PtrSize, "Ptr")
		if (!hSmall || !hLarge) {
			if (hSmall)
				DllCall("DestroyIcon", "Ptr", hSmall)
			if (hLarge)
				DllCall("DestroyIcon", "Ptr", hLarge)
			continue
		}
		; ImageList_AddIcon returns 0-indexed; "Icon" prefix requires 1-indexed.
		smallIdx := DllCall("Comctl32.dll\ImageList_AddIcon", "Ptr", IL.Small, "Ptr", hSmall)
		DllCall("Comctl32.dll\ImageList_AddIcon", "Ptr", IL.Large, "Ptr", hLarge)
		DllCall("DestroyIcon", "Ptr", hSmall)
		DllCall("DestroyIcon", "Ptr", hLarge)

		iconLoadedCount++
		lv_Icons.Add("Icon" . (smallIdx + 1), A_Index)

		if (Mod(iconLoadedCount, 20) = 0)
			SetStatus("Loading icons...", Symbol.Loading, iconLoadedCount . " / " . iconCountTotal)
	}
	lv_Icons.Opt("+Redraw") ; Re-enable redraw
	ApplyViewSettings()
	SetStatus("Ready", Symbol.Success, iconLoadedCount . " icons loaded", State.CurrentDllPath)
}

; Show preview of selected icon
ShowPreview(allowClear := true, *) {
	global State

	selectedRow := lv_Icons.GetNext(0, "Focused")
	if (selectedRow = 0) {
		if (allowClear)
			UpdatePreviewPane("", 0)
		return
	}

	iconNum := GetIconNumberFromRow(selectedRow)
	UpdatePreviewPane(State.CurrentDllPath, iconNum)
}

; Update the entire preview pane (image, labels, code) based on an icon
UpdatePreviewPane(iconPath, iconNum) {
	global PreviewIcon, State, g

	if (iconPath = "" || iconNum = 0) {
		_ClearPreviewBitmap()
		lbl_IconNo.Text := ""
		edt_IconCode.Value := ""
		PreviewIcon := { Path: "", Num: 0 }
		return
	}

	PreviewIcon := { Path: iconPath, Num: iconNum }

	if (!State.Interpolation)
		_RenderPreviewNN(iconPath, iconNum)
	else
		_RenderPreviewHQ(iconPath, iconNum)

	lbl_IconNo.Text := "Icon #" iconNum
	edt_IconCode.Value := 'TraySetIcon("' iconPath '", ' iconNum ')'
}

; Set view mode
SetViewMode(mode) {
	global State
	State.CurrentViewMode := mode
	ApplyViewSettings()
}

; Apply current view settings to ListView
ApplyViewSettings() {
	global mnu_View, g, Txt, State

	isReportView := State.CurrentViewMode < 2
	isLargeIcon := (State.CurrentViewMode = 1) || (State.CurrentViewMode = 3)

	lv_Icons.Opt(isReportView ? "+Report" : "+Icon")
	lv_Icons.SetImageList(isLargeIcon ? IL.Large : IL.Small, isReportView)

	; Update DDL
	ddl_ViewMode.Value := State.CurrentViewMode + 1

	; Update menu checkmarks
	viewItems := [Txt.View0, Txt.View1, Txt.View2, Txt.View3]
	try {
		for item in viewItems
			mnu_View.Uncheck(item)
		mnu_View.Check(viewItems[State.CurrentViewMode + 1])
	}

	; Sync ListView layout for the new view mode without a full GuiSize recalculation
	if (State.LastGuiW > 0) {
		if (isReportView) {
			; Recalculate column width based on current window width
			rightPanelWidth := Layout.Right.W
			if (State.ExportIconSize + 20 > rightPanelWidth)
				rightPanelWidth := State.ExportIconSize + 20
			rightPanelX := State.LastGuiW - rightPanelWidth - Layout.Margin
			newListViewWidth := rightPanelX - Layout.Middle.X - Layout.Middle.Gap
			lv_Icons.ModifyCol(1, newListViewWidth - 20)
		} else {
			SendMessage(0x1016, 0, 0, lv_Icons.Hwnd) ; LVM_ARRANGE
		}
	}
}

; Get icon number from ListView row
GetIconNumberFromRow(rowNumber) {
	if (rowNumber = 0)
		return 0

	; Convert extracted text to integer.
	text := lv_Icons.GetText(rowNumber, 1)
	return Integer(text)
}

; Copies the current icon code to the clipboard for use in scripts.
CopyCurrentCode(*) {
	code := edt_IconCode.Value
	if (code = "") {
		ShowTempTooltip("Select an icon first!", Symbol.Warning)
		return
	}

	A_Clipboard := code
	SetStatus("Copied: " code, Symbol.Copy)
	ShowTempTooltip("Code copied!", Symbol.Copy)
}

CopyIconToClipboard() {
	global PreviewIcon, State

	if (PreviewIcon.Path = "") {
		ShowTempTooltip("Select an icon first!", Symbol.Warning)
		return
	}

	pBitmap := 0
	hIcon := 0
	pStream := 0

	try {
		w := State.ExportIconSize
		h := State.ExportIconSize

		if (!State.Interpolation) {
			DllCall("PrivateExtractIcons", "Str", PreviewIcon.Path, "Int", PreviewIcon.Num - 1,
				"Int", w, "Int", h, "Ptr*", &hIcon, "Ptr", 0, "UInt", 1, "UInt", 0)
			if (!hIcon)
				throw Error("Failed to extract icon")
			pBitmap := Gdip._HICONToBitmap(hIcon, w, h, "NearestNeighbor")
		} else {
			pBitmap := Gdip.ExtractScaledIconBitmap(PreviewIcon.Path, PreviewIcon.Num, w, h)
		}
		if (!pBitmap)
			throw Error("HICON to bitmap failed")

		clsid := Buffer(16)
		DllCall("ole32\CLSIDFromString", "WStr", Gdip.Encoders["png"], "Ptr", clsid)
		DllCall("ole32\CreateStreamOnHGlobal", "Ptr", 0, "Int", false, "Ptr*", &pStream)
		DllCall("gdiplus\GdipSaveImageToStream", "Ptr", pBitmap, "Ptr", pStream, "Ptr", clsid, "Ptr", 0)
		DllCall("ole32\GetHGlobalFromStream", "Ptr", pStream, "Ptr*", &hMemPNG := 0)
		; --- CF_DIBV5 (32-bit Alpha) ---
		rect := Buffer(16, 0)
		NumPut("Int", w, rect, 8), NumPut("Int", h, rect, 12)
		bdSize := (A_PtrSize = 8) ? 32 : 24
		bd := Buffer(bdSize, 0)
		DllCall("gdiplus\GdipBitmapLockBits",
			"Ptr", pBitmap, "Ptr", rect, "UInt", 1,   ; ImageLockModeRead = 1
			"Int", 0x26200A, "Ptr", bd)
		pBits := NumGet(bd, 16, "Ptr")
		stride := Abs(NumGet(bd, 8, "Int"))

		dibHdrSize := 124 ; BITMAPV5HEADER
		dibStride := w * 4
		dibDataSize := dibStride * h
		hMemDIB := DllCall("GlobalAlloc", "UInt", 0x42, "UPtr", dibHdrSize + dibDataSize, "Ptr")
		pDIB := DllCall("GlobalLock", "Ptr", hMemDIB, "Ptr")

		DllCall("RtlZeroMemory", "Ptr", pDIB, "UPtr", dibHdrSize + dibDataSize)

		NumPut("UInt", dibHdrSize, pDIB, 0)     ; bV5Size
		NumPut("Int", w, pDIB, 4)     ; bV5Width
		NumPut("Int", -h, pDIB, 8)     ; bV5Height (top-down)
		NumPut("UShort", 1, pDIB, 12)     ; bV5Planes
		NumPut("UShort", 32, pDIB, 14)     ; bV5BitCount
		NumPut("UInt", 3, pDIB, 16)     ; bV5Compression = BI_BITFIELDS
		NumPut("UInt", dibDataSize, pDIB, 20)     ; bV5SizeImage
		NumPut("UInt", 0x00FF0000, pDIB, 40)     ; bV5RedMask
		NumPut("UInt", 0x0000FF00, pDIB, 44)     ; bV5GreenMask
		NumPut("UInt", 0x000000FF, pDIB, 48)     ; bV5BlueMask
		NumPut("UInt", 0xFF000000, pDIB, 52)     ; bV5AlphaMask
		NumPut("UInt", 0x73524742, pDIB, 56)     ; bV5CSType = LCS_sRGB

		; Stride-optimized pixel copy (single memcpy when strides match)
		pPixelDst := pDIB + dibHdrSize
		if (stride = dibStride) {
			DllCall("RtlMoveMemory", "Ptr", pPixelDst, "Ptr", pBits, "UPtr", dibStride * h)
		} else {
			loop h {
				srcOff := (A_Index - 1) * stride
				dstOff := (A_Index - 1) * dibStride
				DllCall("RtlMoveMemory", "Ptr", pPixelDst + dstOff,
					"Ptr", pBits + srcOff, "UPtr", dibStride)
			}
		}
		DllCall("GlobalUnlock", "Ptr", hMemDIB)

		; Unlock GDI+ bits before creating HBITMAP
		DllCall("gdiplus\GdipBitmapUnlockBits", "Ptr", pBitmap, "Ptr", bd)

		; --- CF_DIB (24-bit) ---
		; GdipCreateHBITMAPFromBitmap composites alpha onto a solid background.
		; GetDIBits then extracts 24-bit bottom-up data for legacy compatibility.
		dibHdrSize24 := 40
		dibStride24 := (w * 3 + 3) & ~3
		dibDataSize24 := dibStride24 * h

		hBmpTemp := 0
		hDC24 := 0
		try {
			DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "Ptr*", &hBmpTemp, "UInt", 0xFF000000)

			hMemDIB24 := DllCall("GlobalAlloc", "UInt", 0x42, "UPtr", dibHdrSize24 + dibDataSize24, "Ptr")
			pDIB24 := DllCall("GlobalLock", "Ptr", hMemDIB24, "Ptr")
			DllCall("RtlZeroMemory", "Ptr", pDIB24, "UPtr", dibHdrSize24 + dibDataSize24)
			NumPut("UInt", dibHdrSize24, pDIB24, 0)
			NumPut("Int", w, pDIB24, 4)
			NumPut("Int", h, pDIB24, 8)
			NumPut("UShort", 1, pDIB24, 12)
			NumPut("UShort", 24, pDIB24, 14)
			NumPut("UInt", 0, pDIB24, 16)
			NumPut("UInt", dibDataSize24, pDIB24, 20)

			hDC24 := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
			DllCall("GetDIBits", "Ptr", hDC24, "Ptr", hBmpTemp, "UInt", 0, "UInt", h,
				"Ptr", pDIB24 + dibHdrSize24, "Ptr", pDIB24, "UInt", 0)
			DllCall("GlobalUnlock", "Ptr", hMemDIB24)
		} finally {
			if (hDC24)
				DllCall("DeleteDC", "Ptr", hDC24)
			if (hBmpTemp)
				DllCall("DeleteObject", "Ptr", hBmpTemp)
		}
		fmtPNG := DllCall("RegisterClipboardFormat", "Str", "PNG", "UInt")
		fmtDIBV5 := 17 ; CF_DIBV5
		fmtDIB := 8  ; CF_DIB
		DllCall("OpenClipboard", "Ptr", 0)
		DllCall("EmptyClipboard")
		DllCall("SetClipboardData", "UInt", fmtPNG, "Ptr", hMemPNG)
		DllCall("SetClipboardData", "UInt", fmtDIBV5, "Ptr", hMemDIB)
		DllCall("SetClipboardData", "UInt", fmtDIB, "Ptr", hMemDIB24)
		DllCall("CloseClipboard")
	}
	finally {
		if (pBitmap)
			DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
		if (pStream)
			ObjRelease(pStream)
		if (hIcon)
			DllCall("DestroyIcon", "Ptr", hIcon)
	}

	SetStatus("Icon copied to clipboard!", Symbol.Copy)
	ShowTempTooltip("Icon copied!", Symbol.Copy)
}

SaveIconToFile(*) {
	global PreviewIcon, g, State

	if (PreviewIcon.Path = "") {
		ShowTempTooltip("Select an icon first!", Symbol.Warning)
		return
	}

	iconPath := PreviewIcon.Path
	iconNum := PreviewIcon.Num

	SplitPath(iconPath, &fileName)
	safeFileName := StrReplace(fileName, ".", "_")
	defaultSaveName := safeFileName "_Icon_" iconNum

	saved := SaveFile([g.Hwnd, "Save Icon"]
		, defaultSaveName
		, { ICO: "*.ico`n", PNG: "*.png", JPEG: "*.jpg", BMP: "*.bmp" }
		, ""
		, 0x6)

	if !saved
		return
	savePath := saved.FileFullPath

	hIcon := 0
	try {
		if (!State.Interpolation) {
			iconIndex := iconNum - 1
			DllCall("PrivateExtractIcons", "Str", iconPath, "Int", iconIndex,
				"Int", State.ExportIconSize, "Int", State.ExportIconSize, "Ptr*", &hIcon, "Ptr", 0, "UInt", 1, "UInt",
				0)
			if (!hIcon)
				throw Error("Failed to extract icon")
			Gdip.SaveHICONToFile(hIcon, savePath, State.ExportIconSize, "NearestNeighbor")
		} else {
			pBitmapSave := Gdip.ExtractScaledIconBitmap(iconPath, iconNum, State.ExportIconSize, State.ExportIconSize)
			try {
				if (!pBitmapSave)
					throw Error("Failed to create bitmap for save")
				Gdip.SaveBitmapToFile(pBitmapSave, savePath)
			} finally {
				if (pBitmapSave)
					DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmapSave)
			}
		}

		SetStatus("Saved: " savePath, Symbol.Success)
		ShowTempTooltip("Icon saved!", Symbol.Success)
	} catch as err {
		MsgBox("Error saving image: " err.Message, "Error", "Icon!")
	} finally {
		if (hIcon)
			DllCall("DestroyIcon", "Ptr", hIcon)
	}
}

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

ShowFileContextMenu(GuiCtrlObj, Item, IsRightClick, X, Y) {
	if !IsRightClick
		return

	if (Item > 0) {
		GuiCtrlObj.Choose(Item) ; Select the right-clicked item
		mnu_FileContext_Item.Show(X, Y)
	} else {
		mnu_FileContext_Empty.Show(X, Y)
	}
}

ShowFavContextMenu(LV, Item, IsRightClick, X, Y) {
	if !IsRightClick
		return

	if (Item > 0) {
		LV.Modify(Item, "Select Focus Vis") ; Select and focus the item
		ShowFavoritePreview()               ; Ensure PreviewIcon is updated
		mnu_FavContext_Item.Show(X, Y)
	} else {
		mnu_FavContext_Empty.Show(X, Y)
	}
}

OpenFileLocation(*) {
	global PreviewIcon, g

	filePath := ""
	if (IsSet(PreviewIcon) && PreviewIcon.Path != "")
		filePath := PreviewIcon.Path
	else if (lst_Files.Value != 0) {
		try {
			txt := lst_Files.Text
			filePath := IsObject(txt) ? txt[1] : txt
		}
	}

	if (filePath != "" && FileExist(filePath))
		Run('explorer.exe /select,"' filePath '"')
}

GoToSource(*) {
	global PreviewIcon, g, FilePathMap

	if (!IsSet(PreviewIcon) || PreviewIcon.Path == "")
		return

	targetPath := PreviewIcon.Path
	targetIcon := PreviewIcon.Num

	; O(1) existence check before iterating
	if (!FilePathMap.Has(targetPath)) {
		ShowTempTooltip("Source file not found in list.", Symbol.Warning)
		return
	}

	; File exists in map, find its index in ListBox
	items := ControlGetItems(lst_Files)
	for index, item in items {
		if (item = targetPath) {
			lst_Files.Choose(0) ; clear previous multi-selection
			lst_Files.Choose(index)
			LoadIcons(true)
			break
		}
	}

	lv_Icons.Modify(0, "-Select")
	loop lv_Icons.GetCount() {
		if (GetIconNumberFromRow(A_Index) = targetIcon) {
			lv_Icons.Modify(A_Index, "Select Focus Vis")
			lv_Icons.Focus()
			break
		}
	}
}

ShowFileProperties(*) {
	global PreviewIcon, g, lst_Files

	filePath := ""
	if (IsSet(PreviewIcon) && PreviewIcon.Path != "")
		filePath := PreviewIcon.Path
	else if (lst_Files.Value != 0) {
		try {
			txt := lst_Files.Text
			filePath := IsObject(txt) ? txt[1] : txt
		}
	}

	if (filePath == "" || !FileExist(filePath))
		return

	try {
		Run('properties "' filePath '"')
		if WinWait("ahk_class #32770", , 2)
			WinSetAlwaysOnTop(1, "A")
	}
}

TestIcon(*) {
	global g, PreviewIcon, State, g_hWndIconSmall, g_hWndIconBig

	testPath := ""
	testNum := 0

	if (IsSet(PreviewIcon) && PreviewIcon.Path != "") {
		testPath := PreviewIcon.Path
		testNum := PreviewIcon.Num
	} else {
		focusedRow := lv_Icons.GetNext(0, "Focused")
		if (focusedRow > 0) {
			testNum := GetIconNumberFromRow(focusedRow)
			testPath := State.CurrentDllPath
		}
	}

	if (testPath == "") {
		MsgBox("Please select an icon from the list to test.", "Warning", "Icon!")
		return
	}

	try {
		TraySetIcon(testPath, testNum)
		hIconSmall := LoadPicture(testPath, "Icon" . testNum . " w" . Metric.SmallIcon.W . " h" . Metric.SmallIcon.H, &
			isIcon)
		if (hIconSmall) {
			SendMessage(0x80, 0, hIconSmall, g.Hwnd) ; ICON_SMALL — window makes its own copy
			if (g_hWndIconSmall)
				DllCall("DestroyIcon", "Ptr", g_hWndIconSmall)
			g_hWndIconSmall := hIconSmall
		}
		hIconBig := LoadPicture(testPath, "Icon" . testNum . " w" . Metric.LargeIcon.W . " h" . Metric.LargeIcon.H, &
			isIcon)
		if (hIconBig) {
			SendMessage(0x80, 1, hIconBig, g.Hwnd)   ; ICON_BIG — window makes its own copy
			if (g_hWndIconBig)
				DllCall("DestroyIcon", "Ptr", g_hWndIconBig)
			g_hWndIconBig := hIconBig
		}
		ShowTempTooltip("Application icon updated", Symbol.Success, 2000)
	} catch as err {
		MsgBox("Error changing icon: " err.Message, "Error", "Icon!")
	}
}

AddToFavorites(*) {
	global State, g

	selectedRow := lv_Icons.GetNext(0, "Focused")
	if (selectedRow = 0) {
		ShowTempTooltip("Select an icon first!", Symbol.Warning)
		return
	}

	iconNum := GetIconNumberFromRow(selectedRow)
	filePath := State.CurrentDllPath

	loop lv_Favorites.GetCount() {
		if (lv_Favorites.GetText(A_Index, 2) = String(iconNum) && lv_Favorites.GetText(A_Index, 4) = filePath) {
			ShowTempTooltip("Icon already in favorites!", Symbol.Warning)
			return
		}
	}

	State.FavoritesListChanged := true

	iconIndex := IL_Add(IL.Favorites, filePath, iconNum)
	SplitPath(filePath, &fileName)
	lv_Favorites.Add("Icon" . iconIndex, "", iconNum, fileName, filePath) ; filePath to 4th column

	SetStatus("Added to favorites: " fileName " " iconNum, Symbol.Star)
	ShowTempTooltip("Added to favorites!", Symbol.Star)
}

RemoveFromFavorites(*) {
	global State, g

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

	DeleteRowsReverse(lv_Favorites, selectedRows)

	State.FavoritesListChanged := true
	ShowTempTooltip("Favorite(s) removed.", Symbol.Remove)
}

ClearFavorites(*) {
	global State, g
	if (lv_Favorites.GetCount() = 0)
		return

	if (MsgBox("Are you sure you want to clear all favorites?", "Clear Favorites", "YesNo Icon?") = "No")
		return

	SendMessage(0x000B, 0, 0, lv_Favorites.Hwnd) ; WM_SETREDRAW off
	lv_Favorites.Delete()
	try IL_Destroy(IL.Favorites)
	IL.Favorites := CreateImageList(Metric.SmallIcon.W, Metric.SmallIcon.H)
	lv_Favorites.SetImageList(IL.Favorites)
	State.FavoritesListChanged := true
	SendMessage(0x000B, 1, 0, lv_Favorites.Hwnd) ; WM_SETREDRAW on
	DllCall("InvalidateRect", "Ptr", lv_Favorites.Hwnd, "Ptr", 0, "Int", true)

	SetStatus("Favorites cleared", Symbol.Trash)
	ShowTempTooltip("Favorites cleared", Symbol.Trash)
}

LoadFavorites() {
	global State, g

	SendMessage(0x000B, 0, 0, lv_Favorites.Hwnd) ; WM_SETREDRAW off
	lv_Favorites.Delete()
	try IL_Destroy(IL.Favorites)
	IL.Favorites := CreateImageList(Metric.SmallIcon.W, Metric.SmallIcon.H)
	lv_Favorites.SetImageList(IL.Favorites)

	if FileExist(App.FavoritesFile) {
		loop read, App.FavoritesFile {
			favString := A_LoopReadLine
			if (Trim(favString) = "")
				continue

			parts := StrSplit(favString, "|")
			if (parts.Length = 2 && FileExist(parts[1])) {
				filePath := parts[1]
				iconNum := Integer(parts[2])

				iconIndex := IL_Add(IL.Favorites, filePath, iconNum)
				SplitPath(filePath, &fileName)
				lv_Favorites.Add("Icon" . iconIndex, "", iconNum, fileName, filePath)
			}
		}
	}
	SendMessage(0x000B, 1, 0, lv_Favorites.Hwnd) ; WM_SETREDRAW on
	DllCall("InvalidateRect", "Ptr", lv_Favorites.Hwnd, "Ptr", 0, "Int", true)
	State.FavoritesListChanged := false
}

SaveFavorites() {
	global State

	if !State.FavoritesListChanged && FileExist(App.FavoritesFile)
		return

	try {
		f := FileOpen(App.FavoritesFile, "w")
		try {
			loop lv_Favorites.GetCount() {
				iconNum := lv_Favorites.GetText(A_Index, 2)
				filePath := lv_Favorites.GetText(A_Index, 4)
				f.WriteLine(filePath "|" iconNum)
			}
		} finally {
			f.Close()
		}
		State.FavoritesListChanged := false
	}
}

ShowFavoritePreview(allowClear := true, *) {

	selected := lv_Favorites.GetNext(0, "Focused")
	if (selected = 0) {
		if (allowClear)
			UpdatePreviewPane("", 0)
		return
	}

	iconNum := Integer(lv_Favorites.GetText(selected, 2))
	filePath := lv_Favorites.GetText(selected, 4)

	UpdatePreviewPane(filePath, iconNum)
}

AddCustomFile(*) {
	selectedFiles := FileSelect("M3", , "Select Icon Source", "Source Files (*.dll; *.exe; *.ico; *.lnk)")
	if (Type(selectedFiles) = "Array")
		AddFilesToList(selectedFiles)
	else if (selectedFiles != "")
		AddFilesToList([selectedFiles])
}

AddFilesToList(FileArray) {
	global Symbol, State, FilePathMap
	addedCount := 0
	lastAddedPath := ""
	SendMessage(0x000B, 0, 0, lst_Files.Hwnd) ; WM_SETREDRAW off
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
		FilePathMap[filePathToAdd] := true
		addedCount++
		lastAddedPath := filePathToAdd
		State.dllFileListChanged := true
	}
	SendMessage(0x000B, 1, 0, lst_Files.Hwnd) ; WM_SETREDRAW on
	DllCall("InvalidateRect", "Ptr", lst_Files.Hwnd, "Ptr", 0, "Int", true)
	if (addedCount > 0) {
		statusMsg := (addedCount = 1) ? "File added" : addedCount " files added"
		SetStatus(statusMsg, Symbol.Success, , lastAddedPath)
	} else if (FileArray.Length > 0) {
		SetStatus("No files added (duplicate or unsupported)", Symbol.Info)
	}
}

RemoveCustomFile(*) {
	global State, g, FilePathMap
	selectedIndices := lst_Files.Value ; Returns an array of indices in v2 multi-select ListBox
	if (selectedIndices.Length = 0) {
		MsgBox("Please select file(s) to remove.", "Info", "Icon!")
		return
	}

	; Remove from FilePathMap before deleting from ListBox
	items := ControlGetItems(lst_Files)
	for idx in selectedIndices {
		if (idx <= items.Length)
			FilePathMap.Delete(items[idx])
	}

	; Delete from bottom to top to prevent index shifting
	DeleteRowsReverse(lst_Files, selectedIndices)

	State.dllFileListChanged := true
	SetStatus("File(s) removed", Symbol.Remove)
}

ClearFileList(*) {
	global State, g, FilePathMap
	if (ControlGetItems(lst_Files).Length == 0)
		return

	if (MsgBox("Are you sure you want to clear the file list?", "Clear List", "YesNo Icon?") = "No")
		return

	SendMessage(0x000B, 0, 0, lst_Files.Hwnd) ; WM_SETREDRAW off
	lst_Files.Delete()
	lv_Icons.Delete()
	FilePathMap.Clear()
	State.CurrentDllPath := ""
	State.dllFileListChanged := true
	SendMessage(0x000B, 1, 0, lst_Files.Hwnd) ; WM_SETREDRAW on
	DllCall("InvalidateRect", "Ptr", lst_Files.Hwnd, "Ptr", 0, "Int", true)
	SetStatus("List cleared", Symbol.Trash)
}

CloseApplication(*) {
	global g_hPreviewBitmap, g_hWndIconSmall, g_hWndIconBig
	SaveFileList()
	SaveFavorites()
	; Clean up preview HBITMAP
	if (g_hPreviewBitmap) {
		DllCall("DeleteObject", "Ptr", g_hPreviewBitmap)
		g_hPreviewBitmap := 0
	}
	; Clean up window icon handles
	if (g_hWndIconSmall) {
		DllCall("DestroyIcon", "Ptr", g_hWndIconSmall)
		g_hWndIconSmall := 0
	}
	if (g_hWndIconBig) {
		DllCall("DestroyIcon", "Ptr", g_hWndIconBig)
		g_hWndIconBig := 0
	}
	; Destroy ImageLists to release GDI resources
	if (IL.Small)
		try DllCall("Comctl32.dll\ImageList_Destroy", "Ptr", IL.Small)
	if (IL.Large)
		try DllCall("Comctl32.dll\ImageList_Destroy", "Ptr", IL.Large)
	if (IL.Favorites)
		try DllCall("Comctl32.dll\ImageList_Destroy", "Ptr", IL.Favorites)

	Gdip.Shutdown()
	ExitApp()
}

SaveFileList() {
	global State, g

	if !State.dllFileListChanged && FileExist(App.DllListFile)
		return

	try {
		f := FileOpen(App.DllListFile, "w")
		try {
			for item in ControlGetItems(lst_Files)
				f.WriteLine(item)
		} finally {
			f.Close()
		}
		State.dllFileListChanged := false
	}
}

IsFileInList(filePath) {
	global FilePathMap
	return FilePathMap.Has(filePath)
}

OnDropFiles(GuiObj, GuiCtrlObj, FileArray, X, Y) {
	AddFilesToList(FileArray)
}

GuiSize(GuiObj, MinMax, Width, Height) {
	if (MinMax = -1)
		return

	global State

	State.LastGuiW := Width
	State.LastGuiH := Height

	GuiRedraw(GuiObj, 0) ; Disable redraw

	; Left panel
	lst_Files.Move(, , , Height - Layout.Left.BottomOffset)

	; Right Panel
	rightPanelWidth := Layout.Right.W
	if (State.ExportIconSize + 20 > rightPanelWidth)
		rightPanelWidth := State.ExportIconSize + 20
	rightPanelX := Width - rightPanelWidth - Layout.Margin

	lbl_PreviewSize.Move(rightPanelX, 10, Layout.Right.LabelW)
	ddl_ExportSize.Move(rightPanelX + Layout.Right.LabelW + 5, 8, Layout.Right.SizeW)

	lbl_IconNo.Move(rightPanelX, 35, 80)
	pic_Preview.Move(rightPanelX, 55, State.ExportIconSize, State.ExportIconSize)
	previewBottom := 55 + State.ExportIconSize
	cb_Interpolation.Move(rightPanelX, previewBottom + 5, rightPanelWidth)
	buttonsY := previewBottom + Layout.Right.BtnH
	btn_CopyImage.Move(rightPanelX, buttonsY)
	btn_SaveImage.Move(rightPanelX + Layout.Right.BtnW + 5, buttonsY)

	iconCodeY := buttonsY + 40
	edt_IconCode.Move(rightPanelX, iconCodeY, rightPanelWidth, 50)

	copyCodeY := iconCodeY + 55
	btn_CopyCode.Move(rightPanelX, copyCodeY)
	btn_Test.Move(rightPanelX + Layout.Right.TestBtnXOffset, copyCodeY)

	favoritesLabelY := copyCodeY + 45
	lbl_Favorites.Move(rightPanelX, favoritesLabelY)

	favButtonsY := favoritesLabelY + 25
	btn_FavAdd.Move(rightPanelX, favButtonsY, Layout.Right.FavBtnW)
	btn_FavRemove.Move(rightPanelX + Layout.Right.FavBtnW + 5, favButtonsY, Layout.Right.FavBtnW)
	btn_FavClear.Move(rightPanelX + (Layout.Right.FavBtnW + 5) * 2, favButtonsY, Layout.Right.FavBtnW)

	favoritesListY := favButtonsY + 35
	favListHeight := Height - favoritesListY - 40
	if (favListHeight < 50)
		favListHeight := 50
	lv_Favorites.Move(rightPanelX, favoritesListY, rightPanelWidth, favListHeight)
	lv_Favorites.ModifyCol(3, rightPanelWidth - Layout.Right.FavIconColW - Layout.Right.FavNumColW - Layout.Margin)

	; Middle Panel
	middlePanelStartX := Layout.Middle.X
	middlePanelGapToRight := Layout.Middle.Gap
	newListViewWidth := rightPanelX - middlePanelStartX - middlePanelGapToRight
	lv_Icons.Move(middlePanelStartX, 35, newListViewWidth, Height - 75)

	isReportView := State.CurrentViewMode < 2
	if (isReportView)
		lv_Icons.ModifyCol(1, newListViewWidth - 20)

	ddl_ViewMode.Move(middlePanelStartX + Layout.Middle.DdlXOffset, 5)

	; Bottom status bar
	sb_Status.Move(, Height - 30, Width - 20)

	; Force rearrange icons if in Icon view
	if (State.CurrentViewMode >= 2)
		SendMessage(0x1016, 0, 0, lv_Icons.Hwnd) ; LVM_ARRANGE

	GuiRedraw(GuiObj, 1) ; Enable redraw
	WinRedraw(GuiObj) ; Force redraw to clear artifacts
}

GuiRedraw(GuiObj, redraw) {
	SendMessage(0x000B, redraw, 0, GuiObj.Hwnd) ; WM_SETREDRAW = 1 (On)
}

; --- Helper Functions ---
/**
 * Returns true if the bitmap has any non-transparent pixels.
 */
HasVisiblePixels(pBitmap) {
	if (!pBitmap)
		return false

	DllCall("gdiplus\GdipGetImageWidth", "Ptr", pBitmap, "UInt*", &w := 0)
	DllCall("gdiplus\GdipGetImageHeight", "Ptr", pBitmap, "UInt*", &h := 0)
	if (w = 0 || h = 0)
		return false

	rect := Buffer(16, 0)
	NumPut("Int", 0, rect, 0), NumPut("Int", 0, rect, 4)
	NumPut("Int", w, rect, 8), NumPut("Int", h, rect, 12)

	bdSize := (A_PtrSize = 8) ? 32 : 24
	bd := Buffer(bdSize, 0)
	if (DllCall("gdiplus\GdipBitmapLockBits", "Ptr", pBitmap, "Ptr", rect, "UInt", 1, "Int", 0x26200A, "Ptr", bd) != 0)
		return true ; assume visible if lock fails

	pBits := NumGet(bd, 16, "Ptr")
	stride := Abs(NumGet(bd, 8, "Int"))

	hasPixels := false
	try {
		if (stride = w * 4) {
			; Optimized pixel scan for non-zero alpha.
			totalPixels := w * h
			quads := totalPixels >> 2
			loop quads {
				off := (A_Index - 1) * 16
				if (NumGet(pBits, off, "UInt") >> 24)
					|| (NumGet(pBits, off + 4, "UInt") >> 24)
					|| (NumGet(pBits, off + 8, "UInt") >> 24)
					|| (NumGet(pBits, off + 12, "UInt") >> 24) {
					hasPixels := true
					break
				}
			}
			; Check remaining pixels.
			if (!hasPixels) {
				rem := totalPixels - (quads * 4)
				off := quads * 16
				loop rem {
					if (NumGet(pBits, off + (A_Index - 1) * 4, "UInt") >> 24) {
						hasPixels := true
						break
					}
				}
			}
		} else {
			; Stride-aware scan with UInt check
			loop h {
				rowPtr := pBits + (A_Index - 1) * stride
				loop w {
					if (NumGet(rowPtr, (A_Index - 1) * 4, "UInt") >> 24) {
						hasPixels := true
						break 2
					}
				}
			}
		}
	} finally {
		DllCall("gdiplus\GdipBitmapUnlockBits", "Ptr", pBitmap, "Ptr", bd)
	}

	return hasPixels
}

/**
 * Sets a GDI+ bitmap directly onto the Picture control via STM_SETIMAGE.
 */
_SetPreviewFromBitmap(pBitmap) {
	global g_hPreviewBitmap
	if (!pBitmap)
		return false

	bgColor := DllCall("GetSysColor", "Int", 15, "UInt") ; COLOR_BTNFACE
	bgR := bgColor & 0xFF
	bgG := (bgColor >> 8) & 0xFF
	bgB := (bgColor >> 16) & 0xFF
	bgARGB := (0xFF << 24) | (bgR << 16) | (bgG << 8) | bgB

	hBitmapNew := 0
	DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "Ptr*", &hBitmapNew, "UInt", bgARGB)
	if (!hBitmapNew)
		return false

	; Ensure SS_BITMAP style (0x0E) — required for STM_SETIMAGE with IMAGE_BITMAP
	curStyle := DllCall("GetWindowLongPtr", "Ptr", pic_Preview.Hwnd, "Int", -16, "Ptr")
	DllCall("SetWindowLongPtr", "Ptr", pic_Preview.Hwnd, "Int", -16, "Ptr", (curStyle & ~0x1F) | 0x0E)

	; Send bitmap to control (STM_SETIMAGE, IMAGE_BITMAP=0)
	SendMessage(0x0172, 0, hBitmapNew, pic_Preview.Hwnd)

	; Delete previous HBITMAP
	if (g_hPreviewBitmap)
		DllCall("DeleteObject", "Ptr", g_hPreviewBitmap)
	g_hPreviewBitmap := hBitmapNew
	return true
}

/**
 * Clears the preview pane by removing the current bitmap image.
 */
_ClearPreviewBitmap() {
	global g_hPreviewBitmap, pic_Preview

	; Send NULL to clear the displayed image — no window recreation needed.
	SendMessage(0x0172, 0, 0, pic_Preview.Hwnd) ; STM_SETIMAGE, IMAGE_BITMAP=0, hBitmap=NULL

	if (g_hPreviewBitmap) {
		DllCall("DeleteObject", "Ptr", g_hPreviewBitmap)
		g_hPreviewBitmap := 0
	}
}

_RenderPreviewNN(iconPath, iconNum) {
	hIcon := 0
	pBitmap := 0
	ok := false
	try {
		w := State.ExportIconSize
		h := State.ExportIconSize
		DllCall("PrivateExtractIcons", "Str", iconPath, "Int", iconNum - 1,
			"Int", w, "Int", h, "Ptr*", &hIcon, "Ptr", 0, "UInt", 1, "UInt", 0)
		if (!hIcon)
			throw Error("Extract failed")

		pBitmap := Gdip._HICONToBitmap(hIcon, w, h, "NearestNeighbor")
		if (!pBitmap)
			throw Error("Bitmap failed")

		if (!HasVisiblePixels(pBitmap))
			throw Error("Empty icon")

		ok := _SetPreviewFromBitmap(pBitmap)
	} catch {
		ok := false
	} finally {
		if (pBitmap)
			DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
		if (hIcon)
			DllCall("DestroyIcon", "Ptr", hIcon)
	}
	if (!ok)
		_ClearPreviewBitmap()
}

_RenderPreviewHQ(iconPath, iconNum) {
	pBitmap := 0
	ok := false
	try {
		w := State.ExportIconSize
		h := State.ExportIconSize
		pBitmap := Gdip.ExtractScaledIconBitmap(iconPath, iconNum, w, h)
		if (!pBitmap)
			throw Error("ExtractScaledIconBitmap failed")
		if (!HasVisiblePixels(pBitmap))
			throw Error("Empty icon")

		ok := _SetPreviewFromBitmap(pBitmap)
	} catch {
		ok := false
	} finally {
		if (pBitmap)
			DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
	}
	if (!ok)
		_ClearPreviewBitmap()
}

SetInterpolation(isOn) {
	global mnu_Icon, cb_Interpolation
	State.Interpolation := isOn ? 1 : 0
	if (State.Interpolation)
		mnu_Icon.Check(Txt.Interpolation)
	else
		mnu_Icon.Uncheck(Txt.Interpolation)
	cb_Interpolation.Value := State.Interpolation
	if (PreviewIcon.Path != "")
		UpdatePreviewPane(PreviewIcon.Path, PreviewIcon.Num)
}

ToggleInterpolation(*) {
	SetInterpolation(!State.Interpolation)
}

OnInterpolationToggle(*) {
	global cb_Interpolation
	SetInterpolation(cb_Interpolation.Value)
}

SetStatus(message, icon := "", iconInfo := "", pathInfo := "") {
	global Symbol, State, g
	if (icon = "" && IsSet(Symbol))
		icon := Symbol.Info

	prefix := (icon != "") ? icon " " : ""
	sb_Status.SetText(prefix . message, 1)

	if (iconInfo != "")
		sb_Status.SetText(iconInfo, 2)

	if (pathInfo != "") {
		sb_Status.SetText(pathInfo, 3)
	} else if (State.CurrentDllPath != "") {
		sb_Status.SetText(State.CurrentDllPath, 3)
	}
}

ShowTempTooltip(message, icon := "", duration := 1500) {
	global Symbol
	if (icon = "" && IsSet(Symbol))
		icon := Symbol.Info
	prefix := (icon != "") ? icon " " : ""
	ToolTip(prefix . message, , , 1)
	SetTimer(() => ToolTip(, , , 1), -duration)
}

OnExportSizeChange(*) {
	global g, PreviewIcon, State

	size := Integer(ddl_ExportSize.Text)
	if (size < 16)
		size := 16
	State.ExportIconSize := size

	if (State.LastGuiW > 0 && State.LastGuiH > 0)
		GuiSize(g, 0, State.LastGuiW, State.LastGuiH)

	if (PreviewIcon.Path != "")
		UpdatePreviewPane(PreviewIcon.Path, PreviewIcon.Num)
}

OnLvKeyDown(wParam, lParam, msg, hwnd) {
	if (!IsSet(lv_Icons) || !lv_Icons)
		return

	if (hwnd = lv_Icons.Hwnd) {
		dir := (wParam = 39) ? 1 : (wParam = 37) ? -1 : 0
		if (dir) { ; Arrow key pressed
			item := lv_Icons.GetNext(0, "Focused")
			nextItem := item + dir
			if (nextItem > 0 && nextItem <= lv_Icons.GetCount()) {
				lv_Icons.Modify(0, "-Focus -Select")
				lv_Icons.Modify(nextItem, "+Focus +Select +Vis")
				ShowPreview()
				return 0
			}
		}
	}
}

ShowAbout(*) {
	aboutText := App.Name " " App.Version
	aboutText .= "
	(
		`n`nA professional utility to browse, preview, and export icons from DLL, EXE, and ICO files.`n
		Mesut Akcan
		mesutakcan.blogspot.com
		github.com/akcansoft
		youtube.com/mesutakcan
	)"

	MsgBox(aboutText, "About", "Iconi")
}

; Deletes items from a ListView or ListBox in reverse order to maintain index integrity.
DeleteRowsReverse(ctrl, indices) {
	if (!indices || indices.Length = 0)
		return

	loop indices.Length {
		idx := indices[indices.Length - A_Index + 1]
		ctrl.Delete(idx)
	}
}
