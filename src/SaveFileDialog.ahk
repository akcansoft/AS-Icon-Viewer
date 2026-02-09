#Requires AutoHotkey v2

; Custom File Save Dialog function
; --------------------------------
; Creates and manages an advanced file save dialog (IFileSaveDialog).
; This function offers more customization than the standard FileSelect function,
; supporting custom filters, default filenames, and custom places.

; Parameters:
; -----------
; Owner        : The Handle (Hwnd) of the owner window or a Gui object.
; FileName     : Default file name or full path.
; Filter       : File type filters (e.g., { "Text File": "*.txt", "Image": "*.png;*.jpg" }).
; CustomPlaces : Custom folder paths to appear in the left panel (Array or single string).
; Options      : Dialog options (flags for FOS_FORCEFILESYSTEM etc.). Default: 0x6.

; Common flags (can be combined):
; 0x00000002 : FOS_OVERWRITEPROMPT - Prompt before overwriting.
; 0x00000004 : FOS_STRICTFILETYPES - Only allow extensions in filter.
; 0x00000008 : FOS_NOCHANGEDIR - Don't change working directory.
; 0x00000020 : FOS_PICKFOLDERS - Pick folders instead of files.
; 0x00000040 : FOS_FORCEFILESYSTEM - Ensure item is in file system.
; 0x00000800 : FOS_PATHMUSTEXIST - Path must exist.
; 0x00001000 : FOS_FILEMUSTEXIST - File must exist.
; 0x00002000 : FOS_CREATEPROMPT - Prompt to create if doesn't exist.
; 0x02000000 : FOS_DONTADDTORECENT - Don't add to recent list.
; 0x10000000 : FOS_FORCESHOWHIDDEN - Show hidden items.

; Returns:
;---------
; An object { FileFullPath: "...", FilterIndex: 1 } if successful, or false if cancelled.

SaveFile(Owner, FileName := "", Filter := "", CustomPlaces := "", Options := 0x6) {
	; Create IFileSaveDialog COM object
	IFileSaveDialog := ComObject("{C0B4E2F3-BA21-4773-8DBA-335EC946EB8B}", "{84BCCD23-5FDE-4CDB-AEA4-AF64B83D78AB}")

	; Process parameters
	Title := IsObject(Owner) ? String(Owner[2]) : ""
	Owner := IsObject(Owner) ? Owner[1] : (WinExist("ahk_id" . Owner) ? Owner : 0)
	Filter := IsObject(Filter) ? Filter : { All_Files: "*.*" }

	; Temporary variables
	Obj := Map()
	IShellItem := 0
	PIDL := 0

	; ==================== FILE NAME AND DIRECTORY SETUP ====================
	if (FileName != "") {
		if InStr(FileName, "\") {
			; Separate if path contains file
			if !(FileName ~= "\\$") {  ; If not ending with \ (has filename)
				File := ""
				SplitPath FileName, &File, &Directory

				; Set filename
				DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 15 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "WStr", File)
			} else {
				Directory := FileName
			}

			; Directory check - set if exists
			while InStr(Directory, "\") && !DirExist(Directory) {
				; Use negative start position for reverse search with InStr
				Directory := SubStr(Directory, 1, InStr(Directory, "\", , -1) - 1)
			}

			if DirExist(Directory) {
				DllCall("Shell32.dll\SHParseDisplayName", "WStr", Directory, "Ptr", 0, "Ptr*", &PIDL, "UInt", 0, "UInt", 0)
				DllCall("Shell32.dll\SHCreateShellItem", "Ptr", 0, "Ptr", 0, "Ptr", PIDL, "Ptr*", &IShellItem)
				Obj[IShellItem] := PIDL

				; Set initial directory
				DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 12 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "Ptr", IShellItem)
			}
		} else {
			; Only filename provided
			DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 15 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "WStr", FileName)
		}
	}

	; ==================== SETUP FILTERS ====================
	FilterCount := 0
	for key, value in Filter.OwnProps() {
		FilterCount++
	}

	; Allocate memory for COMDLG_FILTERSPEC structure
	FilterSpec := Buffer(2 * FilterCount * A_PtrSize)
	FileTypeIndex := 1
	Index := 1
	FirstFilter := ""
	DefaultExt := ""

	for Description, FileTypes in Filter.OwnProps() {
		; Determine default filter (if `n character exists)
		if (Index = 1)
			FirstFilter := FileTypes
		if InStr(FileTypes, "`n") {
			FileTypeIndex := Index
			DefaultExt := _GetFirstFilterExt(FileTypes)
		}

		; Convert description and file types to UTF-16
		DescBuf := Buffer(StrPut(Trim(Description), "UTF-16") * 2)
		StrPut(Trim(Description), DescBuf, "UTF-16")

		TypesBuf := Buffer(StrPut(Trim(StrReplace(FileTypes, "`n")), "UTF-16") * 2)
		StrPut(Trim(StrReplace(FileTypes, "`n")), TypesBuf, "UTF-16")

		; Store buffer references (prevent memory deletion)
		Obj["desc_" . Index] := DescBuf
		Obj["ft_" . Index] := TypesBuf

		; Place pointers into COMDLG_FILTERSPEC structure
		NumPut("Ptr", DescBuf.Ptr, FilterSpec, A_PtrSize * 2 * (Index - 1))
		NumPut("Ptr", TypesBuf.Ptr, FilterSpec, A_PtrSize * (2 * (Index - 1) + 1))

		Index++
	}

	if (DefaultExt = "" && FirstFilter != "")
		DefaultExt := _GetFirstFilterExt(FirstFilter)

	; Assign filters to dialog
	DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 4 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "UInt", FilterCount, "Ptr", FilterSpec)

	; Set default filter index
	DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 5 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "UInt", FileTypeIndex)

	; Set default extension so the dialog appends it before returning
	if (DefaultExt != "")
		DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 22 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "WStr", DefaultExt)

	; ==================== ADD CUSTOM PLACES ====================
	if IsObject(CustomPlaces) || CustomPlaces == "" {
		CustomPlaces := IsObject(CustomPlaces) ? CustomPlaces : (CustomPlaces == "" ? [] : [CustomPlaces])

		for Directory in CustomPlaces {
			Place := 0  ; FDAP_BOTTOM (default)
			DirPath := IsObject(Directory) ? Directory[1] : Directory
			Place := IsObject(Directory) && Directory.Length > 1 ? Directory[2] : 0

			if DirExist(DirPath) {
				DllCall("Shell32.dll\SHParseDisplayName", "WStr", DirPath, "Ptr", 0, "Ptr*", &PIDL, "UInt", 0, "UInt", 0)
				DllCall("Shell32.dll\SHCreateShellItem", "Ptr", 0, "Ptr", 0, "Ptr", PIDL, "Ptr*", &IShellItem)
				Obj[IShellItem] := PIDL

				; Add custom place
				DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 21 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "Ptr", IShellItem, "UInt", Place)
			}
		}
	}

	; ==================== SET TITLE AND OPTIONS ====================
	if (Title != "") {
		DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 17 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "WStr", Title)
	}

	; Set options
	DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 9 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "UInt", Options)

	; ==================== SHOW DIALOG ====================
	Result := false
	HR := DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 3 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "Ptr", Owner, "UInt")

	if (HR == 0) {  ; S_OK
		; Get selected filter index
		DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 6 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "UInt*", &FileTypeIndex)

		; Get selected file
		if !DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 20 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "Ptr*", &IShellItem) {
			ResultBuf := Buffer(32767 * 2, 0)
			DllCall("Shell32.dll\SHGetIDListFromObject", "Ptr", IShellItem, "Ptr*", &PIDL)
			DllCall("Shell32.dll\SHGetPathFromIDListEx", "Ptr", PIDL, "Ptr", ResultBuf, "UInt", 32767, "UInt", 0)
			Result := StrGet(ResultBuf, "UTF-16")
			Obj[IShellItem] := PIDL

			; Extension is now handled by the dialog via SetDefaultExtension.
		}
	}

	; ==================== CLEANUP ====================
	for key, value in Obj {
		if IsInteger(key) {  ; IShellItem interface pointer
			ObjRelease(key)
			DllCall("Ole32.dll\CoTaskMemFree", "Ptr", value)
		}
	}

	; Return result
	return Result ? { FileFullPath: Result, FilterIndex: FileTypeIndex } : false
}

_GetFirstFilterExt(fileTypes) {
	types := Trim(StrReplace(fileTypes, "`n"))
	if InStr(types, ";")
		types := SubStr(types, 1, InStr(types, ";") - 1)
	types := Trim(types)
	types := StrReplace(types, "*.")
	types := StrReplace(types, "*")
	types := Trim(types)
	return (types = "" || types = ".*") ? "" : types
}
