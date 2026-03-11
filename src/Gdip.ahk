#Requires AutoHotkey v2

; GDI+ wrapper for AutoHotkey v2
; Optimized for icon extraction and alpha preservation.
class Gdip {
	static Token := 0
	static Encoders := Map(
		"png", "{557CF406-1A04-11D3-9A73-0000F81EF32E}",
		"bmp", "{557CF400-1A04-11D3-9A73-0000F81EF32E}",
		"jpg", "{557CF401-1A04-11D3-9A73-0000F81EF32E}",
		"jpeg", "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
	)

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

	/*
	Converts an HICON to a GDI+ bitmap.
	Uses DrawIconEx + DIB for 32-bit icons to preserve alpha.
	Uses GdipCreateBitmapFromHICON for legacy icons to read AND mask.
	*/
	static _HICONToBitmap(hIcon, w, h, interpolationMode := "Lanczos") {
		pBitmap := 0
		hDC := 0
		hBitmap := 0
		hOldBmp := 0

		try {
			; Draw icon onto a 32-bit top-down DIB section
			hDC := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
			bi := Buffer(40, 0)
			NumPut("Int", 40, bi, 0)
			NumPut("Int", w, bi, 4)
			NumPut("Int", -h, bi, 8)  ; negative = top-down
			NumPut("Short", 1, bi, 12)
			NumPut("Short", 32, bi, 14)
			pBits := 0
			hBitmap := DllCall("CreateDIBSection", "Ptr", hDC, "Ptr", bi, "UInt", 0,
				"Ptr*", &pBits, "Ptr", 0, "UInt", 0, "Ptr")
			hOldBmp := DllCall("SelectObject", "Ptr", hDC, "Ptr", hBitmap, "Ptr")
			DllCall("RtlZeroMemory", "Ptr", pBits, "UPtr", w * h * 4)
			DllCall("DrawIconEx", "Ptr", hDC, "Int", 0, "Int", 0, "Ptr", hIcon,
				"Int", w, "Int", h, "UInt", 0, "Ptr", 0, "UInt", 3)
			DllCall("GdiFlush")

			; Detect icon type: modern (alpha) vs legacy.
			isModern := this._IsModernIcon(hIcon, w, h, pBits)

			; Apply interpolation mode override
			if (interpolationMode = "NearestNeighbor" && !isModern) {
				; Force NN path for legacy icon, but apply AND mask for transparency.
				isModern := this._ApplyANDMaskToDIB(pBits, w, h, hIcon)
			}

			PixelFormat32bppPARGB := 0xE200B

			if (isModern) {
				; Modern icon: Copy DIB pixels to PARGB bitmap to avoid edge artifacts.
				DllCall("gdiplus\GdipCreateBitmapFromScan0",
					"Int", w, "Int", h, "Int", 0,
					"Int", PixelFormat32bppPARGB,
					"Ptr", 0, "Ptr*", &pBitmap)
				if (!pBitmap)
					throw Error("GdipCreateBitmapFromScan0 failed")

				rect := Buffer(16, 0)
				NumPut("Int", w, rect, 8), NumPut("Int", h, rect, 12)
				bdSize := (A_PtrSize = 8) ? 32 : 24
				bd := Buffer(bdSize, 0)
				DllCall("gdiplus\GdipBitmapLockBits",
					"Ptr", pBitmap, "Ptr", rect, "UInt", 2,
					"Int", PixelFormat32bppPARGB, "Ptr", bd)
				dstScan0 := NumGet(bd, 16, "Ptr")
				dstStride := Abs(NumGet(bd, 8, "Int"))
				srcStride := w * 4
				if (dstStride = srcStride) {
					; Strides match: copy entire pixel buffer in one call
					DllCall("RtlMoveMemory", "Ptr", dstScan0,
						"Ptr", pBits, "UPtr", srcStride * h)
				} else {
					loop h {
						srcOff := (A_Index - 1) * srcStride
						dstOff := (A_Index - 1) * dstStride
						DllCall("RtlMoveMemory", "Ptr", dstScan0 + dstOff,
							"Ptr", pBits + srcOff, "UPtr", srcStride)
					}
				}
				DllCall("gdiplus\GdipBitmapUnlockBits", "Ptr", pBitmap, "Ptr", bd)
			} else {
				; Legacy icon: GdipCreateBitmapFromHICON handles AND mask transparency.
				DllCall("gdiplus\GdipCreateBitmapFromHICON", "Ptr", hIcon, "Ptr*", &pBitmap)
				if (!pBitmap)
					throw Error("GdipCreateBitmapFromHICON failed")
			}
		} finally {
			if (hOldBmp && hDC)
				DllCall("SelectObject", "Ptr", hDC, "Ptr", hOldBmp)
			if (hBitmap)
				DllCall("DeleteObject", "Ptr", hBitmap)
			if (hDC)
				DllCall("DeleteDC", "Ptr", hDC)
		}

		return pBitmap
	}

	; Applies the AND mask from an HICON onto a 32-bpp top-down DIB buffer.
	; Sets alpha=0 where mask bit is 1 (transparent), alpha=255 otherwise.
	static _ApplyANDMaskToDIB(pBits, w, h, hIcon) {
		if (!pBits || !hIcon || w < 1 || h < 1)
			return false

		ii := Buffer(32, 0)
		if !DllCall("GetIconInfo", "Ptr", hIcon, "Ptr", ii, "Int")
			return false

		hbmMask := NumGet(ii, 16, "Ptr")
		hbmColor := NumGet(ii, 16 + A_PtrSize, "Ptr")

		try {
			if (!hbmMask)
				return false

			bm := Buffer(32, 0)
			DllCall("GetObject", "Ptr", hbmMask, "Int", 32, "Ptr", bm)
			maskW := NumGet(bm, 4, "Int")
			maskH := NumGet(bm, 8, "Int")
			if (maskW < 1 || maskH < 1)
				return false

			; Check for monochrome icons (stacked AND+XOR masks).
			if (!hbmColor && maskH >= h * 2)
				maskH := maskH // 2

			applyW := (w < maskW) ? w : maskW
			applyH := (h < maskH) ? h : maskH

			bmi := Buffer(48, 0) ; BITMAPINFOHEADER + 2 RGBQUADs
			NumPut("UInt", 40, bmi, 0)
			NumPut("Int", maskW, bmi, 4)
			NumPut("Int", -maskH, bmi, 8) ; top-down
			NumPut("UShort", 1, bmi, 12)
			NumPut("UShort", 1, bmi, 14)
			NumPut("UInt", 0, bmi, 16)

			stride := ((maskW + 31) // 32) * 4
			maskBuf := Buffer(stride * maskH, 0)

			hDC := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
			hOld := DllCall("SelectObject", "Ptr", hDC, "Ptr", hbmMask, "Ptr")
			scanLines := DllCall("GetDIBits", "Ptr", hDC, "Ptr", hbmMask,
				"UInt", 0, "UInt", maskH, "Ptr", maskBuf, "Ptr", bmi, "UInt", 0)
			DllCall("SelectObject", "Ptr", hDC, "Ptr", hOld)
			DllCall("DeleteDC", "Ptr", hDC)

			if (scanLines = 0)
				return false

			loop applyH {
				y := A_Index - 1
				rowBase := y * stride
				pRowBits := pBits + (y * w * 4)

				; Process in chunks of 8 pixels to minimize NumGet calls
				loop (applyW // 8) {
					byteIdx := A_Index - 1
					maskByte := NumGet(maskBuf, rowBase + byteIdx, "UChar")
					pPixelAlpha := pRowBits + (byteIdx * 32) + 3

					NumPut("UChar", (maskByte & 0x80) ? 0 : 255, pPixelAlpha)
					NumPut("UChar", (maskByte & 0x40) ? 0 : 255, pPixelAlpha + 4)
					NumPut("UChar", (maskByte & 0x20) ? 0 : 255, pPixelAlpha + 8)
					NumPut("UChar", (maskByte & 0x10) ? 0 : 255, pPixelAlpha + 12)
					NumPut("UChar", (maskByte & 0x08) ? 0 : 255, pPixelAlpha + 16)
					NumPut("UChar", (maskByte & 0x04) ? 0 : 255, pPixelAlpha + 20)
					NumPut("UChar", (maskByte & 0x02) ? 0 : 255, pPixelAlpha + 24)
					NumPut("UChar", (maskByte & 0x01) ? 0 : 255, pPixelAlpha + 28)
				}

				; Handle remaining pixels (< 8)
				rem := Mod(applyW, 8)
				if (rem) {
					startX := applyW - rem
					maskByte := NumGet(maskBuf, rowBase + (startX >> 3), "UChar")
					loop rem {
						bit := 0x80 >> (A_Index - 1)
						NumPut("UChar", (maskByte & bit) ? 0 : 255, pRowBits + (startX + A_Index - 1) * 4 + 3)
					}
				}
			}

			return true
		} finally {
			if (hbmMask)
				DllCall("DeleteObject", "Ptr", hbmMask)
			if (hbmColor)
				DllCall("DeleteObject", "Ptr", hbmColor)
		}
	}

	static ExtractScaledIconBitmap(iconPath, iconNum, targetW, targetH) {
		hIconTest := 0
		hIconSrc := 0
		pBitmapSrc := 0
		pBitmapDst := 0
		pGraphics := 0

		try {
			; ── Step 1: lightweight 32×32 extraction to detect modern vs legacy ──
			DllCall("PrivateExtractIcons", "Str", iconPath, "Int", iconNum - 1,
				"Int", 32, "Int", 32, "Ptr*", &hIconTest, "Ptr", 0, "UInt", 1, "UInt", 0)
			if (!hIconTest)
				throw Error("Failed to extract icon for type detection")

			; Render to a temporary DIB and look for any non-zero alpha byte
			isModern := this._IsModernIcon(hIconTest, 32, 32)
			DllCall("DestroyIcon", "Ptr", hIconTest), hIconTest := 0

			; Choose source size based on icon type (256 for modern, 32 for legacy).
			srcW := isModern ? 256 : 32
			srcH := srcW

			DllCall("PrivateExtractIcons", "Str", iconPath, "Int", iconNum - 1,
				"Int", srcW, "Int", srcH, "Ptr*", &hIconSrc, "Ptr", 0, "UInt", 1, "UInt", 0)
			if (!hIconSrc)
				throw Error("Failed to extract icon at source size " srcW)

			; Convert HICON → GDI+ bitmap at source size
			; For modern icons the DIB/PARGB path is used; for legacy, GdipCreateBitmapFromHICON.
			pBitmapSrc := this._HICONToBitmap(hIconSrc, srcW, srcH, "Lanczos")
			DllCall("DestroyIcon", "Ptr", hIconSrc), hIconSrc := 0
			if (!pBitmapSrc)
				throw Error("Failed to create source bitmap")

			; ── Step 3: return directly if no scaling is needed ──
			if (srcW = targetW && srcH = targetH) {
				result := pBitmapSrc
				pBitmapSrc := 0   ; transfer ownership to caller
				return result
			}

			; Apply HighQualityBicubic scaling.
			DllCall("gdiplus\GdipCreateBitmapFromScan0",
				"Int", targetW, "Int", targetH, "Int", 0, "Int", 0x26200A,  ; PixelFormat32bppARGB
				"Ptr", 0, "Ptr*", &pBitmapDst)
			if (!pBitmapDst)
				throw Error("Failed to create destination bitmap")

			DllCall("gdiplus\GdipGetImageGraphicsContext", "Ptr", pBitmapDst, "Ptr*", &pGraphics)
			DllCall("gdiplus\GdipSetInterpolationMode", "Ptr", pGraphics, "Int", 7) ; HighQualityBicubic
			DllCall("gdiplus\GdipSetPixelOffsetMode", "Ptr", pGraphics, "Int", 2) ; PixelOffsetModeHalf
			DllCall("gdiplus\GdipSetSmoothingMode", "Ptr", pGraphics, "Int", 2) ; AntiAlias
			DllCall("gdiplus\GdipDrawImageRectRect",
				"Ptr", pGraphics, "Ptr", pBitmapSrc,
				"Float", 0.0, "Float", 0.0, "Float", targetW, "Float", targetH,
				"Float", 0.0, "Float", 0.0, "Float", srcW, "Float", srcH,
				"Int", 2, "Ptr", 0, "Ptr", 0, "Ptr", 0)

			DllCall("gdiplus\GdipDeleteGraphics", "Ptr", pGraphics), pGraphics := 0
			DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmapSrc), pBitmapSrc := 0

			result := pBitmapDst
			pBitmapDst := 0     ; transfer ownership to caller
			return result

		} finally {
			if (hIconTest)
				DllCall("DestroyIcon", "Ptr", hIconTest)
			if (hIconSrc)
				DllCall("DestroyIcon", "Ptr", hIconSrc)
			if (pGraphics)
				DllCall("gdiplus\GdipDeleteGraphics", "Ptr", pGraphics)
			if (pBitmapSrc)
				DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmapSrc)
			if (pBitmapDst)
				DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmapDst)
		}
	}

	; Saves a GDI+ bitmap (pBitmap) to a file.
	; Supported formats: PNG, BMP, JPG/JPEG, ICO (PNG-in-ICO with full alpha).
	static SaveBitmapToFile(pBitmap, filePath) {
		if (!pBitmap)
			throw Error("pBitmap is null")

		SplitPath(filePath, , , &ext)
		ext := StrLower(ext)

		if (ext = "ico") {
			; PNG-in-ICO format (full alpha preservation).
			imgW := 0, imgH := 0
			DllCall("gdiplus\GdipGetImageWidth", "Ptr", pBitmap, "UInt*", &imgW)
			DllCall("gdiplus\GdipGetImageHeight", "Ptr", pBitmap, "UInt*", &imgH)
			this._WriteICOFile(filePath, pBitmap, imgW, imgH)
		}
		else if this.Encoders.Has(ext) {
			clsid := Buffer(16)
			DllCall("ole32\CLSIDFromString", "WStr", this.Encoders[ext], "Ptr", clsid)
			status := DllCall("gdiplus\GdipSaveImageToFile",
				"Ptr", pBitmap, "WStr", filePath, "Ptr", clsid, "Ptr", 0, "Int")
			if (status != 0)
				throw Error("GdipSaveImageToFile failed (status " status ")")
		}
		else
			throw Error("Unsupported format: " ext)
	}

	static SaveHICONToFile(hIcon, filePath, iconSize := 0, interpolationMode := "Lanczos") {
		if (!hIcon)
			throw Error("hIcon is null")

		SplitPath(filePath, , , &ext)
		ext := StrLower(ext)

		; Determine icon dimensions
		if (iconSize > 0) {
			w := iconSize
			h := iconSize
		} else {
			ii := Buffer(32, 0)
			DllCall("GetIconInfo", "Ptr", hIcon, "Ptr", ii)
			hBmpMask := NumGet(ii, 16, "Ptr")
			hBmpColor := NumGet(ii, 16 + A_PtrSize, "Ptr")
			bm := Buffer(32, 0)
			DllCall("GetObject", "Ptr", hBmpColor ? hBmpColor : hBmpMask, "Int", 32, "Ptr", bm)
			w := NumGet(bm, 4, "Int")
			h := NumGet(bm, 8, "Int")
			if (hBmpColor)
				DllCall("DeleteObject", "Ptr", hBmpColor)
			if (hBmpMask)
				DllCall("DeleteObject", "Ptr", hBmpMask)

			if (w < 1)
				w := 32
			if (h < 1)
				h := 32
		}

		pBitmap := 0
		try {
			pBitmap := this._HICONToBitmap(hIcon, w, h, interpolationMode)
			if (ext = "ico") {
				this._WriteICOFile(filePath, pBitmap, w, h)
			}
			else if this.Encoders.Has(ext) {
				clsid := Buffer(16)
				DllCall("ole32\CLSIDFromString", "WStr", this.Encoders[ext], "Ptr", clsid)
				status := DllCall("gdiplus\GdipSaveImageToFile",
					"Ptr", pBitmap, "WStr", filePath, "Ptr", clsid, "Ptr", 0, "Int")
				if (status != 0)
					throw Error("GdipSaveImageToFile failed (status " status ")")
			}
			else
				throw Error("Unsupported format: " ext)
		} finally {
			if (pBitmap)
				DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
		}
	}

	static _WriteICOFile(filePath, pBitmap, w, h) {
		pStream := 0
		try {
			pngClsid := Buffer(16)
			DllCall("ole32\CLSIDFromString", "WStr", this.Encoders["png"], "Ptr", pngClsid)
			DllCall("ole32\CreateStreamOnHGlobal", "Ptr", 0, "Int", true, "Ptr*", &pStream)
			DllCall("gdiplus\GdipSaveImageToStream", "Ptr", pBitmap, "Ptr", pStream, "Ptr", pngClsid, "Ptr", 0)

			hMem := 0
			DllCall("ole32\GetHGlobalFromStream", "Ptr", pStream, "Ptr*", &hMem)
			pngSize := DllCall("GlobalSize", "Ptr", hMem, "UPtr")
			pData := DllCall("GlobalLock", "Ptr", hMem, "Ptr")
			pngBuf := Buffer(pngSize)
			DllCall("RtlMoveMemory", "Ptr", pngBuf, "Ptr", pData, "UPtr", pngSize)
			DllCall("GlobalUnlock", "Ptr", hMem)

			imgW := (w >= 1) ? w : 256
			imgH := (h >= 1) ? h : 256

			f := FileOpen(filePath, "w")
			if (!f)
				throw Error("Failed to open file for writing: " filePath)
			try {
				f.WriteUShort(0)                           ; reserved
				f.WriteUShort(1)                           ; type = ICO
				f.WriteUShort(1)                           ; image count
				f.WriteUChar(imgW >= 256 ? 0 : imgW)      ; width  (0 = 256)
				f.WriteUChar(imgH >= 256 ? 0 : imgH)      ; height (0 = 256)
				f.WriteUChar(0)                            ; color count
				f.WriteUChar(0)                            ; reserved
				f.WriteUShort(1)                           ; planes
				f.WriteUShort(32)                          ; bit depth
				f.WriteUInt(pngSize)                       ; image data size
				f.WriteUInt(22)                            ; offset to image data
				f.RawWrite(pngBuf)
			} finally {
				f.Close()
			}
		} finally {
			if (pStream)
				ObjRelease(pStream)
		}
	}

	static _IsModernIcon(hIcon, w, h, pExternalBits := 0) {
		hDC := 0
		hBitmap := 0
		hOldBmp := 0
		pBits := pExternalBits
		isModern := false

		try {
			if (!pBits) {
				hDC := DllCall("CreateCompatibleDC", "Ptr", 0, "Ptr")
				bi := Buffer(40, 0)
				NumPut("Int", 40, bi, 0)
				NumPut("Int", w, bi, 4)
				NumPut("Int", -h, bi, 8)
				NumPut("Short", 1, bi, 12)
				NumPut("Short", 32, bi, 14)
				hBitmap := DllCall("CreateDIBSection", "Ptr", hDC, "Ptr", bi, "UInt", 0,
					"Ptr*", &pBits, "Ptr", 0, "UInt", 0, "Ptr")
				hOldBmp := DllCall("SelectObject", "Ptr", hDC, "Ptr", hBitmap, "Ptr")
				DllCall("RtlZeroMemory", "Ptr", pBits, "UPtr", w * h * 4)
				DllCall("DrawIconEx", "Ptr", hDC, "Int", 0, "Int", 0, "Ptr", hIcon,
					"Int", w, "Int", h, "UInt", 0, "Ptr", 0, "UInt", 3)
				DllCall("GdiFlush")
			}

			; Optimized selection scan for non-zero alpha.
			totalPixels := w * h
			quads := totalPixels >> 2
			loop quads {
				off := (A_Index - 1) * 16
				if (NumGet(pBits, off, "UInt") >> 24)
					|| (NumGet(pBits, off + 4, "UInt") >> 24)
					|| (NumGet(pBits, off + 8, "UInt") >> 24)
					|| (NumGet(pBits, off + 12, "UInt") >> 24) {
					isModern := true
					break
				}
			}
			; Check remaining pixels (0-3)
			if (!isModern) {
				rem := totalPixels - (quads * 4)
				off := quads * 16
				loop rem {
					if (NumGet(pBits, off + (A_Index - 1) * 4, "UInt") >> 24) {
						isModern := true
						break
					}
				}
			}
		} finally {
			if (hOldBmp)
				DllCall("SelectObject", "Ptr", hDC, "Ptr", hOldBmp)
			if (hBitmap)
				DllCall("DeleteObject", "Ptr", hBitmap)
			if (hDC)
				DllCall("DeleteDC", "Ptr", hDC)
		}
		return isModern
	}
}
