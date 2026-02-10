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
		if (!hIcon)
			throw Error("hIcon is null")

		SplitPath(filePath, , , &ext)
		ext := StrLower(ext)

		static encoders := Map(
			"png", "{557CF406-1A04-11D3-9A73-0000F81EF32E}",
			"bmp", "{557CF400-1A04-11D3-9A73-0000F81EF32E}",
			"jpg", "{557CF401-1A04-11D3-9A73-0000F81EF32E}",
			"jpeg", "{557CF401-1A04-11D3-9A73-0000F81EF32E}",
			"ico", "{557CF406-1A04-11D3-9A73-0000F81EF32E}"  ; PNG → ICO wrapper
		)

		pBitmap := 0
		DllCall("gdiplus\GdipCreateBitmapFromHICON", "Ptr", hIcon, "Ptr*", &pBitmap)

		try {
			; Convert premultiplied alpha to straight alpha to avoid black fringe
			pBitmapToSave := pBitmap
			pBitmapARGB := this._UnpremultiplyBitmap(pBitmap)
			if (pBitmapARGB)
				pBitmapToSave := pBitmapARGB

			if (ext = "ico") {
				; Mevcut ICO wrapper mantığın iyi, sadece biraz temizledim
				tempPng := A_Temp "\asiv_temp_" A_TickCount ".png"
				clsid := Buffer(16)
				DllCall("ole32\CLSIDFromString", "WStr", encoders["png"], "Ptr", clsid)
				DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmapToSave, "WStr", tempPng, "Ptr", clsid, "Ptr", 0)

				pngData := FileRead(tempPng, "RAW")
				FileDelete(tempPng)

				w := 0, h := 0
				DllCall("gdiplus\GdipGetImageWidth", "Ptr", pBitmapToSave, "UInt*", &w)
				DllCall("gdiplus\GdipGetImageHeight", "Ptr", pBitmapToSave, "UInt*", &h)
				if (w < 1)
					w := 256
				if (h < 1)
					h := 256
				wByte := (w >= 256) ? 0 : w
				hByte := (h >= 256) ? 0 : h

				f := FileOpen(filePath, "w")
				f.WriteUShort(0)      ; reserved
				f.WriteUShort(1)      ; type = ICO
				f.WriteUShort(1)      ; count
				f.WriteUChar(wByte)   ; width (0 means 256)
				f.WriteUChar(hByte)   ; height (0 means 256)
				f.WriteUChar(0)       ; colors
				f.WriteUChar(0)       ; reserved
				f.WriteUShort(1)      ; planes
				f.WriteUShort(32)     ; bit count
				f.WriteUInt(pngData.Size)
				f.WriteUInt(22)       ; offset
				f.RawWrite(pngData)
				f.Close()
			}
			else if encoders.Has(ext) {
				clsid := Buffer(16)
				DllCall("ole32\CLSIDFromString", "WStr", encoders[ext], "Ptr", clsid)
				DllCall("gdiplus\GdipSaveImageToFile", "Ptr", pBitmapToSave, "WStr", filePath, "Ptr", clsid, "Ptr", 0)
			}
			else
				throw Error("Unsupported format: " ext)
		}
		finally {
			if (pBitmapARGB)
				DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmapARGB)
			if (pBitmap)
				DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
		}
	}

	static _UnpremultiplyBitmap(pBitmap) {
		if (!pBitmap)
			return 0

		PixelFormat32bppPARGB := 0xE200B
		PixelFormat32bppARGB := 0x26200A

		pixelFormat := 0
		if (DllCall("gdiplus\GdipGetImagePixelFormat", "Ptr", pBitmap, "Int*", &pixelFormat) != 0)
			return 0
		if (pixelFormat != PixelFormat32bppPARGB)
			return 0

		width := 0, height := 0
		DllCall("gdiplus\GdipGetImageWidth", "Ptr", pBitmap, "UInt*", &width)
		DllCall("gdiplus\GdipGetImageHeight", "Ptr", pBitmap, "UInt*", &height)
		if (width = 0 || height = 0)
			return 0

		pBitmapARGB := 0
		if (DllCall("gdiplus\GdipCreateBitmapFromScan0", "Int", width, "Int", height, "Int", 0
			, "Int", PixelFormat32bppARGB, "Ptr", 0, "Ptr*", &pBitmapARGB) != 0)
			return 0

		rect := Buffer(16, 0)
		NumPut("Int", 0, rect, 0), NumPut("Int", 0, rect, 4)
		NumPut("Int", width, rect, 8), NumPut("Int", height, rect, 12)

		bdSize := (A_PtrSize = 8) ? 32 : 24
		bdSrc := Buffer(bdSize, 0)
		bdDst := Buffer(bdSize, 0)
		lockedSrc := false
		lockedDst := false

		try {
			if (DllCall("gdiplus\GdipBitmapLockBits", "Ptr", pBitmap, "Ptr", rect, "UInt", 1
				, "Int", PixelFormat32bppPARGB, "Ptr", bdSrc) != 0)
				return 0
			lockedSrc := true

			if (DllCall("gdiplus\GdipBitmapLockBits", "Ptr", pBitmapARGB, "Ptr", rect, "UInt", 2
				, "Int", PixelFormat32bppARGB, "Ptr", bdDst) != 0)
				return 0
			lockedDst := true

			srcStride := NumGet(bdSrc, 8, "Int")
			dstStride := NumGet(bdDst, 8, "Int")
			srcScan0 := NumGet(bdSrc, 16, "Ptr")
			dstScan0 := NumGet(bdDst, 16, "Ptr")

			if (srcStride < 0) {
				srcStride := -srcStride
				srcBase := srcScan0 + (height - 1) * srcStride
				srcDir := -1
			} else {
				srcBase := srcScan0
				srcDir := 1
			}

			if (dstStride < 0) {
				dstStride := -dstStride
				dstBase := dstScan0 + (height - 1) * dstStride
				dstDir := -1
			} else {
				dstBase := dstScan0
				dstDir := 1
			}

			loop height {
				y := A_Index - 1
				srcRow := srcBase + y * srcStride * srcDir
				dstRow := dstBase + y * dstStride * dstDir

				loop width {
					x := A_Index - 1
					pixel := NumGet(srcRow + x * 4, 0, "UInt")
					a := (pixel >> 24) & 0xFF
					if (a) {
						r := (pixel >> 16) & 0xFF
						g := (pixel >> 8) & 0xFF
						b := pixel & 0xFF

						r := (r * 255 + (a // 2)) // a
						g := (g * 255 + (a // 2)) // a
						b := (b * 255 + (a // 2)) // a

						if (r > 255)
							r := 255
						if (g > 255)
							g := 255
						if (b > 255)
							b := 255
					} else {
						r := 0, g := 0, b := 0
					}

					NumPut("UInt", (a << 24) | (r << 16) | (g << 8) | b, dstRow + x * 4)
				}
			}
		}
		finally {
			if (lockedSrc)
				DllCall("gdiplus\GdipBitmapUnlockBits", "Ptr", pBitmap, "Ptr", bdSrc)
			if (lockedDst)
				DllCall("gdiplus\GdipBitmapUnlockBits", "Ptr", pBitmapARGB, "Ptr", bdDst)
		}

		return pBitmapARGB
	}

}