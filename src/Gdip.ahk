#Requires AutoHotkey v2

; GDI+ wrapper for AutoHotkey v2
class Gdip {
	static Token := 0

	static Startup() {
		if (this.Token)
			return
		DllCall("LoadLibrary", "Str", "gdiplus", "Ptr")
		si := Buffer(24, 0), NumPut("UInt", 1, si)
		DllCall("gdiplus\GdiplusStartup", "Ptr*", &token := 0, "Ptr", si, "Ptr", 0)
		this.Token := token
	}

	static Shutdown() {
		if (this.Token) {
			DllCall("gdiplus\GdiplusShutdown", "Ptr", this.Token)
			this.Token := 0
		}
	}

	static SaveHICONToFile(hIcon, filePath) {
		SplitPath(filePath, , , &ext)
		ext := StrLower(ext)

		clsids := Map(
			"png", "{557CF406-1A04-11D3-9A73-0000F81EF32E}",
			"bmp", "{557CF400-1A04-11D3-9A73-0000F81EF32E}",
			"jpg", "{557CF401-1A04-11D3-9A73-0000F81EF32E}",
			"jpeg", "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
		)

		pBitmap := 0
		if (DllCall("gdiplus\GdipCreateBitmapFromHICON", "Ptr", hIcon, "Ptr*", &pBitmap) != 0)
			throw Error("Failed to create GDI+ bitmap from HICON.")

		try {
			if (ext = "ico") {
				tempFile := A_Temp "\temp_icon_" A_TickCount ".png"
				DllCall("ole32\CLSIDFromString", "Str", clsids["png"], "Ptr", Encoder := Buffer(16))
				DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "WStr", tempFile, "Ptr", Encoder, "Ptr", 0)
				try {
					pngData := FileRead(tempFile, "RAW")
					f := FileOpen(filePath, "w")
					f.WriteUShort(0), f.WriteUShort(1), f.WriteUShort(1)
					f.WriteUChar(128), f.WriteUChar(128), f.WriteUChar(0), f.WriteUChar(0)
					f.WriteUShort(1), f.WriteUShort(32)
					f.WriteUInt(pngData.Size), f.WriteUInt(22)
					f.RawWrite(pngData)
					f.Close()
				} finally {
					if FileExist(tempFile)
						FileDelete(tempFile)
				}
			} else if clsids.Has(ext) {
				DllCall("ole32\CLSIDFromString", "Str", clsids[ext], "Ptr", Encoder := Buffer(16))
				DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmap, "WStr", filePath, "Ptr", Encoder, "Ptr", 0)
			}
		} finally {
			if (pBitmap)
				DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
		}
	}
}