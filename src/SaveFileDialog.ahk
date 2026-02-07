; Custom File Save Dialog function
SaveFile(Owner, FileName := "", Filter := "", CustomPlaces := "", Options := 0x6) {
    ; Create IFileSaveDialog COM object
    IFileSaveDialog := ComObject("{C0B4E2F3-BA21-4773-8DBA-335EC946EB8B}", "{84BCCD23-5FDE-4CDB-AEA4-AF64B83D78AB}")
    
    ; Process parameters
    Title := IsObject(Owner) ? String(Owner[2]) : ""
    Owner := IsObject(Owner) ? Owner[1] : (WinExist("ahk_id" . Owner) ? Owner : 0)
    Filter := IsObject(Filter) ? Filter : {All_Files: "*.*"}
    
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
                Directory := SubStr(Directory, 1, InStr(Directory, "\",, -1) - 1)
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
    
    for Description, FileTypes in Filter.OwnProps() {
        ; Determine default filter (if `n character exists)
        if InStr(FileTypes, "`n") {
            FileTypeIndex := Index
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
    
    ; Assign filters to dialog
    DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 4 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "UInt", FilterCount, "Ptr", FilterSpec)
    
    ; Set default filter index
    DllCall(NumGet(NumGet(IFileSaveDialog.Ptr, "Ptr") + 5 * A_PtrSize, "Ptr"), "Ptr", IFileSaveDialog, "UInt", FileTypeIndex)
    
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
            
            ; ==================== AUTOMATIC EXTENSION ADDING ====================
            ; If file has no extension, add based on selected filter
            SplitPath Result, , , &Ext
            if (Ext == "") {
                ; Find selected filter
                CurrentIndex := 1
                for Description, FileTypes in Filter.OwnProps() {
                    if (CurrentIndex == FileTypeIndex) {
                        ; Get first extension from filter (*.txt -> txt)
                        FilterExt := Trim(StrReplace(FileTypes, "`n"))
                        ; Get first one if multiple extensions (*.doc;*.docx -> *.doc)
                        if InStr(FilterExt, ";") {
                            FilterExt := SubStr(FilterExt, 1, InStr(FilterExt, ";") - 1)
                        }
                        ; Remove *. and get only extension
                        FilterExt := StrReplace(FilterExt, "*.")
                        FilterExt := StrReplace(FilterExt, "*")
                        FilterExt := Trim(FilterExt)
                        
                        ; Add if valid extension exists
                        if (FilterExt != "" && FilterExt != ".*") {
                            Result .= "." . FilterExt
                        }
                        break
                    }
                    CurrentIndex++
                }
            }
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
    return Result ? {FileFullPath: Result, FilterIndex: FileTypeIndex} : false
}
