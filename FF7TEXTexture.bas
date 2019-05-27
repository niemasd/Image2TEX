Attribute VB_Name = "FF7TEXTexture"
Option Explicit
Type TEXTexture
    tex_file As String
    hdc As Long
    hbmp As Long
    
    'TEX file format by Mirex and Aali
    'http://wiki.qhimm.com/FF7/TEX_format
    Version As Long             'Must be 1 for FF7 to load it
    unk1 As Long
    ColorKeyFlag As Long        'Set to 1 to enable the transparent color
    unk2 As Long
    unk3 As Long
    MinimumBitsPerColor As Long 'D3D driver uses these to determine which texture
                                'format to convert to on load
    MaximumBitsPerColor As Long
    MinimumAlphaBits As Long
    MaximumAlphaBits As Long
    MinimumBitsPerPixel As Long
    MaximumBitsPerPixel As Long
    unk4 As Long
    NumPalletes As Long
    NumColorsPerPallete As Long
    BitDepth As Long
    width As Long
    height As Long
    BytesPerRow As Long         'Rarelly used. Usually assumed to be BytesperPixel*Width
    unk5 As Long
    PalleteFlag As Long
    BitsPerIndex As Long
    IndexedTo8bitsFlag As Long  'Never used in FF7
    PalleteSize As Long         'Must be NumPalletes*NumColorsPerPallete
    NumColorsPerPallete2 As Long 'Can be the same or 0. Doesn't really matter
    RuntimeData As Long         'Placeholder for some information. Doesn't matter
    BitsPerPixel As Long
    BytesPerPixel As Long       'Should be trusted over BitsPerPixel
    'Pixel format (all 0 for palletized images)
    NumRedBits As Long
    NumGreenBits As Long
    NumBlueBits As Long
    NumAlphaBits As Long
    RedBitMask As Long
    GreenBitMask As Long
    BlueBitMask As Long
    AlphaBitMask As Long
    RedShift As Long
    GreenShift As Long
    BlueShift As Long
    AlphaShift As Long
    'The components values are computed by the following expresion:
    '(((value & mask) >> shift) * 255) / max
    Red8 As Long                'Allways 8
    Green8 As Long              'Allways 8
    Blue8 As Long               'Allways 8
    Alpha8 As Long              'Allways 8
    RedMax As Long
    GreenMax As Long
    BlueMax As Long
    AlphaMax As Long
    'End of Pixel format
    ColorKeyArrayFlag As Long
    RuntimeData2 As Long
    ReferenceAlpha As Long
    unk6 As Long
    unk7 As Long
    RuntimeDataPalleteIndex As Long
    RuntimeData3 As Long
    RuntimeData4 As Long
    unk8 As Long
    unk9 As Long
    unk10 As Long
    unk11 As Long
    Pallete() As Byte          'Always in 32-bit BGRA format
    PixelData() As Byte         'Width * Height * BytesPerPixel. Either indices or raw
                                'data following the specified format
    ColorKeyData() As Byte      'NumPalletes * 1 bytes
End Type
Function ReadTEXTexture(ByRef Texture As TEXTexture, ByVal fileName As String) As Integer
    Dim NFile As Integer
    Dim offBitmap As Long
    
    On Error GoTo errorH
    
    If FileExist(fileName) Then
        NFile = FreeFile

        Open fileName For Binary As NFile
        
        With Texture
            .tex_file = fileName
            
            Get NFile, 1, .Version
            Get NFile, 1 + &H4, .unk1
            Get NFile, 1 + &H8, .ColorKeyFlag
            Get NFile, 1 + &HC, .unk2
            Get NFile, 1 + &H10, .unk3
            Get NFile, 1 + &H14, .MinimumBitsPerColor
            Get NFile, 1 + &H18, .MaximumBitsPerColor
            Get NFile, 1 + &H1C, .MinimumAlphaBits
            Get NFile, 1 + &H20, .MaximumAlphaBits
            Get NFile, 1 + &H24, .MinimumBitsPerPixel
            Get NFile, 1 + &H28, .MaximumBitsPerPixel
            Get NFile, 1 + &H2C, .unk4
            Get NFile, 1 + &H30, .NumPalletes
            Get NFile, 1 + &H34, .NumColorsPerPallete
            Get NFile, 1 + &H38, .BitDepth
            Get NFile, 1 + &H3C, .width
            Get NFile, 1 + &H40, .height
            Get NFile, 1 + &H44, .BytesPerRow
            Get NFile, 1 + &H48, .unk5
            Get NFile, 1 + &H4C, .PalleteFlag
            Get NFile, 1 + &H50, .BitsPerIndex
            Get NFile, 1 + &H54, .IndexedTo8bitsFlag
            Get NFile, 1 + &H58, .PalleteSize
            Get NFile, 1 + &H5C, .NumColorsPerPallete2
            Get NFile, 1 + &H60, .RuntimeData
            Get NFile, 1 + &H64, .BitsPerPixel
            Get NFile, 1 + &H68, .BytesPerPixel
            
            Get NFile, 1 + &H6C, .NumRedBits
            Get NFile, 1 + &H70, .NumGreenBits
            Get NFile, 1 + &H74, .NumBlueBits
            Get NFile, 1 + &H78, .NumAlphaBits
            Get NFile, 1 + &H7C, .RedBitMask
            Get NFile, 1 + &H80, .GreenBitMask
            Get NFile, 1 + &H84, .BlueBitMask
            Get NFile, 1 + &H88, .AlphaBitMask
            Get NFile, 1 + &H8C, .RedShift
            Get NFile, 1 + &H90, .GreenShift
            Get NFile, 1 + &H94, .BlueShift
            Get NFile, 1 + &H98, .AlphaShift
            Get NFile, 1 + &H9C, .Red8
            Get NFile, 1 + &HA0, .Green8
            Get NFile, 1 + &HA4, .Blue8
            Get NFile, 1 + &HA8, .Alpha8
            Get NFile, 1 + &HAC, .RedMax
            Get NFile, 1 + &HB0, .GreenMax
            Get NFile, 1 + &HB4, .BlueMax
            Get NFile, 1 + &HB8, .AlphaMax
            
            Get NFile, 1 + &HBC, .ColorKeyArrayFlag
            Get NFile, 1 + &HC0, .RuntimeData2
            Get NFile, 1 + &HC4, .ReferenceAlpha
            Get NFile, 1 + &HC8, .unk6
            Get NFile, 1 + &HCC, .unk7
            Get NFile, 1 + &HD0, .RuntimeDataPalleteIndex
            Get NFile, 1 + &HD4, .RuntimeData3
            Get NFile, 1 + &HD8, .RuntimeData4
            Get NFile, 1 + &HDC, .unk8
            Get NFile, 1 + &HE0, .unk9
            Get NFile, 1 + &HE4, .unk10
            Get NFile, 1 + &HE8, .unk11
            
            offBitmap = 1 + &HEC
            If .PalleteFlag = 1 Then
                ReDim .Pallete(.PalleteSize * 4 - 1)
                Get NFile, 1 + &HEC, .Pallete
                offBitmap = offBitmap + .PalleteSize * 4
            End If
                
            ReDim .PixelData(.width * .height * .BytesPerPixel - 1)
            Get NFile, offBitmap, .PixelData
            
            If .ColorKeyArrayFlag = 1 Then
                ReDim .ColorKeyData(.NumPalletes - 1)
                Get NFile, offBitmap + .width * .height * .BytesPerPixel, .ColorKeyData
            End If
        End With
        Close NFile
        ReadTEXTexture = 0
    Else
        Debug.Print "TEX file not found!!!"
        MsgBox "TEX file " + fileName + " not found!!!", vbOKOnly, "Error reading"
        ReadTEXTexture = -1
    End If
    Exit Function
errorH:
    MsgBox "Error opening " + fileName, vbCritical, "TEX Error " + Str$(Err)
End Function
Sub WriteTEXTexture(ByRef Texture As TEXTexture, ByVal fileName As String)
    Dim NFile As Integer
    Dim offBitmap As Long
    
    On Error GoTo errorH
    
    NFile = FreeFile
    'Truncate file first
    Open fileName For Output As NFile
    Close NFile
        
    Open fileName For Binary As NFile
    
    With Texture
        Put NFile, 1, .Version
        Put NFile, 1 + &H4, .unk1
        Put NFile, 1 + &H8, .ColorKeyFlag
        Put NFile, 1 + &HC, .unk2
        Put NFile, 1 + &H10, .unk3
        Put NFile, 1 + &H14, .MinimumBitsPerColor
        Put NFile, 1 + &H18, .MaximumBitsPerColor
        Put NFile, 1 + &H1C, .MinimumAlphaBits
        Put NFile, 1 + &H20, .MaximumAlphaBits
        Put NFile, 1 + &H24, .MinimumBitsPerPixel
        Put NFile, 1 + &H28, .MaximumBitsPerPixel
        Put NFile, 1 + &H2C, .unk4
        Put NFile, 1 + &H30, .NumPalletes
        Put NFile, 1 + &H34, .NumColorsPerPallete
        Put NFile, 1 + &H38, .BitDepth
        Put NFile, 1 + &H3C, .width
        Put NFile, 1 + &H40, .height
        Put NFile, 1 + &H44, .BytesPerRow
        Put NFile, 1 + &H48, .unk5
        Put NFile, 1 + &H4C, .PalleteFlag
        Put NFile, 1 + &H50, .BitsPerIndex
        Put NFile, 1 + &H54, .IndexedTo8bitsFlag
        Put NFile, 1 + &H58, .PalleteSize
        Put NFile, 1 + &H5C, .NumColorsPerPallete2
        Put NFile, 1 + &H60, .RuntimeData
        Put NFile, 1 + &H64, .BitsPerPixel
        Put NFile, 1 + &H68, .BytesPerPixel
        
        Put NFile, 1 + &H6C, .NumRedBits
        Put NFile, 1 + &H70, .NumGreenBits
        Put NFile, 1 + &H74, .NumBlueBits
        Put NFile, 1 + &H78, .NumAlphaBits
        Put NFile, 1 + &H7C, .RedBitMask
        Put NFile, 1 + &H80, .GreenBitMask
        Put NFile, 1 + &H84, .BlueBitMask
        Put NFile, 1 + &H88, .AlphaBitMask
        Put NFile, 1 + &H8C, .RedShift
        Put NFile, 1 + &H90, .GreenShift
        Put NFile, 1 + &H94, .BlueShift
        Put NFile, 1 + &H98, .AlphaShift
        Put NFile, 1 + &H9C, .Red8
        Put NFile, 1 + &HA0, .Green8
        Put NFile, 1 + &HA4, .Blue8
        Put NFile, 1 + &HA8, .Alpha8
        Put NFile, 1 + &HAC, .RedMax
        Put NFile, 1 + &HB0, .GreenMax
        Put NFile, 1 + &HB4, .BlueMax
        Put NFile, 1 + &HB8, .AlphaMax
        
        Put NFile, 1 + &HBC, .ColorKeyArrayFlag
        Put NFile, 1 + &HC0, .RuntimeData2
        Put NFile, 1 + &HC4, .ReferenceAlpha
        Put NFile, 1 + &HC8, .unk6
        Put NFile, 1 + &HCC, .unk7
        Put NFile, 1 + &HD0, .RuntimeDataPalleteIndex
        Put NFile, 1 + &HD4, .RuntimeData3
        Put NFile, 1 + &HD8, .RuntimeData4
        Put NFile, 1 + &HDC, .unk8
        Put NFile, 1 + &HE0, .unk9
        Put NFile, 1 + &HE4, .unk10
        Put NFile, 1 + &HE8, .unk11
        offBitmap = 1 + &HEC
        If .PalleteFlag = 1 Then
            Put NFile, 1 + &HEC, .Pallete
            offBitmap = offBitmap + .PalleteSize * 4
        End If
            
        Put NFile, offBitmap, .PixelData
        
        If .ColorKeyArrayFlag = 1 Then
            Put NFile, offBitmap + .width * .height * .BytesPerPixel, .ColorKeyData
        End If
    End With
    Close NFile
    Exit Sub
errorH:
    MsgBox "Error writting " + fileName, vbCritical, "TEX Error " + Str$(Err)
End Sub
Sub GetTEXTextureFromBitmap(ByRef tex_out As TEXTexture, ByRef hdc As Long, ByVal hbmp As Long)
    Dim ci As Long
    Dim li As Long
    Dim si As Long
    Dim ti As Long
    Dim PI As Integer
    Dim aux_val As Byte
    Dim pal_size As Long
    Dim Bits As Integer
    Dim TexBitmapSize As Long
    Dim LineLength As Long
    Dim LineLengthBytes As Long
    Dim LinePad As Long
    Dim LinePadUseful As Long
    Dim LinePadBytes As Long
    Dim Shift As Integer
    Dim mask As Integer
    Dim parts As Integer
    Dim parts_left As Integer
    Dim line_end As Long
    
    Dim PicInfo As BITMAPINFO
    Dim PicData() As Byte
    
    GetAllBitmapData hdc, hbmp, PicData, PicInfo
    
    Bits = PicInfo.bmiHeader.biBitCount
    pal_size = IIf(Bits <= 8, 2 ^ Bits, 0)
    With tex_out
        .Version = 1
        .unk1 = 0
        .ColorKeyFlag = 1
        .unk2 = 1
        .unk3 = 5
        .MinimumBitsPerColor = Bits
        .MaximumBitsPerColor = 8
        .MinimumAlphaBits = 0
        .MaximumAlphaBits = 8
        .MinimumBitsPerPixel = 8
        .MaximumBitsPerPixel = 32
        .unk4 = 0
        .NumPalletes = IIf(pal_size > 0, 1, 0)
        .NumColorsPerPallete = pal_size
        .BitDepth = Bits
        .width = PicInfo.bmiHeader.biWidth
        .height = PicInfo.bmiHeader.biHeight
        .BytesPerRow = IIf(Bits < 8, .width, (Bits * .width) / 8)
        .unk5 = 0
        .PalleteFlag = IIf(Bits <= 8, 1, 0)
        .BitsPerIndex = IIf(Bits <= 8, 8, 0)
        .IndexedTo8bitsFlag = 0
        .PalleteSize = pal_size
        .NumColorsPerPallete2 = pal_size
        .RuntimeData = 19752016
        .BitsPerPixel = Bits
        .BytesPerPixel = IIf(Bits < 8, 1, Bits / 8)
        .Red8 = 8
        .Green8 = 8
        .Blue8 = 8
        .Alpha8 = 8
        Select Case Bits
            Case 16:
                .NumRedBits = 5
                .NumGreenBits = 5
                .NumBlueBits = 5
                .NumAlphaBits = 0
                .RedBitMask = &H7E00
                .GreenBitMask = &H3E0
                .BlueBitMask = &H1F
                .AlphaBitMask = 0
                .RedShift = 10
                .GreenShift = 5
                .BlueShift = 0
                .AlphaShift = 0
            Case 24:
                .NumRedBits = 8
                .NumGreenBits = 8
                .NumBlueBits = 8
                .NumAlphaBits = 0
                .RedBitMask = &HFF0000
                .GreenBitMask = &HFF00
                .BlueBitMask = &HFF
                .AlphaBitMask = 0
                .RedShift = 16
                .GreenShift = 8
                .BlueShift = 0
                .AlphaShift = 0
            Case 32:
                .NumRedBits = 8
                .NumGreenBits = 8
                .NumBlueBits = 8
                .NumAlphaBits = 8
                .RedBitMask = &HFF0000
                .GreenBitMask = &HFF00
                .BlueBitMask = &HFF
                .AlphaBitMask = &HFF000000
                .RedShift = 16
                .GreenShift = 8
                .BlueShift = 0
                .AlphaShift = 24
        End Select
        .RedMax = 2 ^ .NumRedBits - 1
        .GreenMax = 2 ^ .NumGreenBits - 1
        .BlueMax = 2 ^ .NumBlueBits - 1
        .AlphaMax = 2 ^ .NumAlphaBits - 1
        .ColorKeyArrayFlag = 0
        .RuntimeData2 = 0
        .ReferenceAlpha = 255
        .unk6 = 4
        .unk7 = 1
        .RuntimeDataPalleteIndex = 0
        .RuntimeData3 = 34546076
        .RuntimeData4 = 0
        .unk8 = 0
        .unk9 = 480
        .unk10 = 320
        .unk11 = 512
        
        LineLength = .width * .BitsPerPixel
        LinePad = IIf(LineLength Mod 32 = 0, 0, 32 * (LineLength \ 32 + 1) - 8 * (LineLength \ 8))
        LinePadUseful = IIf(LinePad = 0, 0, LineLength + -8 * (LineLength \ 8))
        LinePadBytes = IIf(LinePad > 0 And LinePad < 8, 1, LinePad \ 8)
        LineLengthBytes = LineLength \ 8 + LinePadBytes
        TexBitmapSize = .width * .height * .BytesPerPixel - 1
        ReDim .PixelData(TexBitmapSize)
        
        If Bits = 1 Or Bits = 4 Then
            ti = 0
            Shift = 2 ^ Bits
            mask = Shift - 1
            parts = 8 \ Bits - 1
            parts_left = LinePadUseful \ Bits - 1
            For li = .height - 2 To 0 Step -1
                line_end = (li + 1) * LineLengthBytes - LinePadBytes - 1
                For si = li * LineLengthBytes To line_end
                    aux_val = PicData(si)
                    For PI = 0 To parts
                        .PixelData(ti + parts - PI) = aux_val And mask
                        aux_val = aux_val \ Shift
                    Next PI
                    ti = ti + parts + 1
                Next si
                If LinePad > 0 Then
                    aux_val = PicData(si)
                    For PI = 0 To parts_left
                        .PixelData(ti) = aux_val And mask
                        aux_val = aux_val \ Shift
                    Next PI
                    ti = ti + parts_left + 1
                End If
            Next li
        Else
            For li = 0 To .height - 1
                CopyMemory .PixelData(li * LineLength \ 8), _
                    PicData(((.height - 1) - li) * LineLengthBytes), _
                    LineLength \ 8
            Next li
        End If
        
        If .PalleteFlag = 1 Then
            ReDim .Pallete(4 * .NumColorsPerPallete - 1)
            CopyMemory .Pallete(0), PicInfo.bmiColors(0), 4 * .NumColorsPerPallete
            For ci = 0 To .NumColorsPerPallete - 1
                .Pallete(ci * 4 + 3) = &HFF
            Next ci
        End If
    End With
End Sub
'Create the bitmap object to blit it to any HDC
Sub LoadBitmapFromTEXTexture(ByRef tex_in As TEXTexture)
    Dim ci As Long
    Dim li As Long
    Dim si As Long
    Dim ti As Long
    Dim PI As Integer
    Dim aux_val As Byte
    Dim pal_size As Long
    Dim BMPSizeBytes As Long
    Dim LineLength As Long
    Dim LineLengthBytes As Long
    Dim LinePad As Long
    Dim LinePadUseful As Long
    Dim LinePadBytes As Long
    Dim Shift As Integer
    Dim mask As Integer
    Dim parts As Integer
    Dim parts_left As Integer
    Dim line_end As Long
    
    Dim PicInfo As BITMAPINFO
    Dim PicData() As Byte
    
    With PicInfo.bmiHeader
        .biSize = 40
        .biWidth = tex_in.width
        .biHeight = tex_in.height
        .biPlanes = 1
        If tex_in.PalleteFlag = 1 Then
            .biBitCount = Log(tex_in.PalleteSize) / Log(2)
        Else
            .biBitCount = tex_in.BitDepth
        End If
        .biCompression = BI_RGB
        
        LineLength = .biWidth * .biBitCount
        LinePad = IIf(LineLength Mod 32 = 0, 0, 32 * (LineLength \ 32 + 1) - 8 * (LineLength \ 8))
        LinePadUseful = IIf(LinePad = 0, 0, LineLength - 8 * (LineLength \ 8))
        LinePadBytes = IIf(LinePad > 0 And LinePad < 8, 1, LinePad \ 8)
        LineLengthBytes = LineLength \ 8 + LinePadBytes
        BMPSizeBytes = LineLengthBytes * .biHeight
        
        .biSizeImage = BMPSizeBytes
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biClrUsed = IIf(tex_in.ColorKeyFlag = 1, tex_in.PalleteSize, 0)
        .biClrImportant = .biClrUsed
    
    
        If (PicInfo.bmiHeader.biBitCount <= 8) Then
            Dim aux As Integer
            aux = tex_in.PalleteSize * 4
            
            CopyMemory PicInfo.bmiColors(0), _
                        tex_in.Pallete(0), _
                        aux
        End If

        ReDim PicData(BMPSizeBytes - 1)
        
        If .biBitCount = 1 Or .biBitCount = 4 Then
            si = 0
            Shift = 2 ^ .biBitCount
            mask = Shift - 1
            parts = 8 \ .biBitCount - 1
            parts_left = LinePadUseful \ .biBitCount - 1
            For li = .biHeight - 2 To 0 Step -1
                line_end = (li + 1) * LineLengthBytes - LinePadBytes - 1
                For ti = li * LineLengthBytes To line_end
                    aux_val = 0
                    For PI = 0 To parts
                        aux_val = aux_val * Shift Or tex_in.PixelData(si)
                        si = si + 1
                    Next PI
                    PicData(ti) = aux_val
                Next ti
                If LinePad > 0 Then
                    aux_val = 0
                    For PI = 0 To parts_left
                        aux_val = aux_val * Shift Or tex_in.PixelData(si)
                    Next PI
                    si = IIf(parts_left > -1, -1, 0)
                    ti = ti + LinePadBytes
                End If
            Next li
        Else
            For li = 0 To .biHeight - 1
                CopyMemory PicData(((.biHeight - 1) - li) * LineLengthBytes), _
                            tex_in.PixelData(li * LineLength \ 8), _
                            LineLength \ 8
            Next li
        End If
    End With
    tex_in.hdc = CreateCompatibleDC(0)
    tex_in.hbmp = CreateDIBSection(tex_in.hdc, PicInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    SelectObject tex_in.hdc, tex_in.hbmp
    SetDIBits tex_in.hdc, tex_in.hbmp, 0, PicInfo.bmiHeader.biHeight, PicData(0), PicInfo, DIB_RGB_COLORS
End Sub
Sub GetTEXTextureFromPictureBox(ByRef tex_out As TEXTexture, ByRef hdc As Long, ByVal width As Integer, ByVal height As Integer)
    Dim x As Integer, y As Integer
    Dim bi As Long
    Dim col32 As Long
    Dim pal_size As Long
    Dim Bits As Integer
    Dim offBit As Integer
    
    Bits = 24
    pal_size = 0
    With tex_out
        .Version = 1
        .unk1 = 0
        .ColorKeyFlag = 1
        .unk2 = 1
        .unk3 = 5
        .MinimumBitsPerColor = Bits
        .MaximumBitsPerColor = 8
        .MinimumAlphaBits = 0
        .MaximumAlphaBits = 8
        .MinimumBitsPerPixel = 8
        .MaximumBitsPerPixel = 32
        .unk4 = 0
        .NumPalletes = 0
        .NumColorsPerPallete = pal_size
        .BitDepth = Bits
        .width = width
        .height = height
        .BytesPerRow = IIf(Bits < 8, .width, (Bits * .width) / 8)
        .unk5 = 0
        .PalleteFlag = 0
        .BitsPerIndex = 0
        .IndexedTo8bitsFlag = 0
        .PalleteSize = pal_size
        .NumColorsPerPallete2 = pal_size
        .RuntimeData = 19752016
        .BitsPerPixel = Bits
        .BytesPerPixel = IIf(Bits < 8, 1, Bits / 8)
        .Red8 = 8
        .Green8 = 8
        .Blue8 = 8
        .Alpha8 = 8
        Select Case Bits
            Case 16:
                .NumRedBits = 5
                .NumGreenBits = 5
                .NumBlueBits = 5
                .NumAlphaBits = 1
                .RedBitMask = &H7E00
                .GreenBitMask = &H3E0
                .BlueBitMask = &H1F
                .AlphaBitMask = &H8000
                .RedShift = 10
                .GreenShift = 5
                .BlueShift = 0
                .AlphaShift = 15
            Case 24:
                .NumRedBits = 8
                .NumGreenBits = 8
                .NumBlueBits = 8
                .NumAlphaBits = 0
                .RedBitMask = &HFF0000
                .GreenBitMask = &HFF00
                .BlueBitMask = &HFF
                .AlphaBitMask = 0
                .RedShift = 16
                .GreenShift = 8
                .BlueShift = 0
                .AlphaShift = 0
            Case 32:
                .NumRedBits = 8
                .NumGreenBits = 8
                .NumBlueBits = 8
                .NumAlphaBits = 8
                .RedBitMask = &HFF0000
                .GreenBitMask = &HFF00
                .BlueBitMask = &HFF
                .AlphaBitMask = &HFF000000
                .RedShift = 16
                .GreenShift = 8
                .BlueShift = 0
                .AlphaShift = 24
        End Select
        .RedMax = 2 ^ .NumRedBits - 1
        .GreenMax = 2 ^ .NumGreenBits - 1
        .BlueMax = 2 ^ .NumBlueBits - 1
        .AlphaMax = 2 ^ .NumAlphaBits - 1
        .ColorKeyArrayFlag = 0
        .RuntimeData2 = 0
        .ReferenceAlpha = 255
        .unk6 = 4
        .unk7 = 1
        .RuntimeDataPalleteIndex = 0
        .RuntimeData3 = 34546076
        .RuntimeData4 = 0
        .unk8 = 0
        .unk9 = 480
        .unk10 = 320
        .unk11 = 512
        ReDim .PixelData(.width * .height * .BytesPerPixel - 1)
        bi = 0
                
        For y = 0 To height - 1
            For x = 0 To width - 1
                col32 = GetPixel(hdc, x, y) 'pic.point(X, Y)
                .PixelData(bi) = (col32 \ 2 ^ 16) And 255
                .PixelData(bi + 1) = (col32 \ 2 ^ 8) And 255
                .PixelData(bi + 2) = col32 And 255
                bi = bi + 3
            Next x
        Next y
    End With
End Sub
Public Function FileExist(asPath As String) As Boolean
    If UCase(Dir(asPath)) = UCase(TrimPath(asPath)) Then
        FileExist = True
    Else
        FileExist = False
    End If
End Function
Public Function TrimPath(ByVal asPath As String) As String
    If Len(asPath) = 0 Then Exit Function
    Dim x As Integer

    Do
        x = InStr(asPath, "\")
        If x = 0 Then Exit Do
        asPath = Right(asPath, Len(asPath) - x)
    Loop
    TrimPath = asPath
End Function
