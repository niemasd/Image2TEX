VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Image2TEX 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image2TEX 0.4 - by Borde (Special thanks to Aali)"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   Icon            =   "Image2TEX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton MassConvertCommand 
      Caption         =   "Mass convert"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox TransparentFlagCheck 
      Caption         =   "Color 0 as transparent"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   0
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton SaveTextureButton 
      Caption         =   "Save Texture"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton OpenTextureButton 
      Caption         =   "Open Image"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox TextureViewer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2655
      Left            =   0
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   477
      TabIndex        =   0
      Top             =   480
      Width           =   7215
   End
End
Attribute VB_Name = "Image2TEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MinWidth As Long
Dim MinHeight As Long
Const BorderWidth As Long = 4
Dim hDC_Tex As Long
Dim hbmp_Tex As Long
Dim Width_Tex As Long
Dim Height_Tex As Long
Dim TexLoadedQ As Boolean
Private Sub Form_Load()
    Dim aux_h As Long
    MinWidth = TransparentFlagCheck.Left + TransparentFlagCheck.width
    aux_h = Me.height
    Me.height = 0
    MinHeight = Me.height
    Me.height = aux_h
    TexLoadedQ = False
End Sub

Private Sub MassConvertCommand_Click()
    Dim pattern As String
    Dim tex As TEXTexture
    Dim Bmp As BMPTexture
    Dim hdc As Long
    Dim hbmp As Long
    
    Dim fileName As String
    Dim filename_out As String
    Dim filenames_list() As String
    Dim path As String
    Dim num_entries As Integer
    Dim num_files As Integer
    Dim first_index As Integer
    Dim output_tex_file As Boolean
    Dim file_count As Integer
    
    On Error GoTo hand_op_mass
    pattern = "Any Image file(*.bmp;*.jpg;*.gif;*.ico;*.rle;*.wmf;*.emf)|*.bmp;*.jpg;*.gif;*.ico;*.wmf;*.emf|TEX texture|*.tex"
    CommonDialog1.Filter = pattern
    CommonDialog1.CancelError = True
    CommonDialog1.MaxFileSize = 32000
    CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly    'Allow multiple files to be selected
    CommonDialog1.ShowOpen 'Display the Open File Common Dialog
    
        
    If (CommonDialog1.fileName <> "") Then
        file_count = 0
        
        filenames_list = Split(CommonDialog1.fileName, Chr(0))
        num_entries = UBound(filenames_list)
        
        If num_entries > 0 Then
            If InStr(filenames_list(1), "\") > 0 Then
                'According to the documentation this used to work this way, but it doesn't on my system.
                path = ""
                num_files = num_entries
                first_index = 0
            Else
                path = filenames_list(0) + "\"
                num_files = num_entries - 1
                first_index = 1
            End If
        Else
            num_files = 0
            first_index = 0
            path = ""
        End If
        
        Do
            fileName = path + filenames_list(first_index + num_files)
            
            On Error GoTo hand_op_mass
            If (UCase(Right(fileName, 3)) = "TEX") Then
                ReadTEXTexture tex, fileName
                LoadBitmapFromTEXTexture tex
                hDC_Tex = tex.hdc
                hbmp_Tex = tex.hbmp
                Width_Tex = tex.width
                Height_Tex = tex.height
                TextureViewer.AutoSize = False
                TextureViewer.AutoRedraw = False
                TextureViewer.width = (tex.width + BorderWidth) * 15
                TextureViewer.height = (tex.height + BorderWidth) * 15
                TextureViewer.Picture = Nothing
                TexLoadedQ = True
                TransparentFlagCheck.Value = IIf(tex.ColorKeyFlag = 1, vbChecked, vbUnchecked)
                filename_out = Left(fileName, Len(fileName) - 3) + "BMP"
                output_tex_file = False
            Else
                Set TextureViewer.Picture = LoadPicture(fileName)
                TextureViewer.AutoRedraw = True
                TextureViewer.AutoSize = True
                If (TexLoadedQ) Then
                    DeleteDC hDC_Tex
                    DeleteObject hbmp_Tex
                    TexLoadedQ = False
                End If
                filename_out = Left(fileName, Len(fileName) - 3) + "TEX"
                output_tex_file = True
            End If
        
            Image2TEX.width = IIf(TextureViewer.width < MinWidth, _
                                MinWidth, TextureViewer.width)
            Dim h As Long
            h = MinHeight + TextureViewer.Top + TextureViewer.height
            Image2TEX.height = h
            
            If TexLoadedQ Then
                hdc = hDC_Tex
                hbmp = hbmp_Tex
            Else
                hdc = TextureViewer.hdc
                hbmp = TextureViewer.Picture.handle
            End If
             
            On Error GoTo hand_sav_mass
            If output_tex_file Then
                GetTEXTextureFromBitmap tex, hdc, hbmp
                tex.ColorKeyFlag = IIf(TransparentFlagCheck.Value = vbChecked, 1, 0)
                WriteTEXTexture tex, filename_out
            Else
                GetBMPTextureFromBitmap Bmp, hdc, hbmp
                WriteBMPTexture Bmp, filename_out
            End If
            
            num_files = num_files - 1
            file_count = file_count + 1
        Loop Until num_files < 0
        Erase filenames_list
        
        MsgBox "Processed " + Str$(file_count) + " files succesfully.", vbOKOnly, "Done"
    End If
    Exit Sub
hand_op_mass:
    If Err <> 32755 Then
        MsgBox "Error" + Str$(Err) + " loading " + fileName, vbOKOnly, "Unknow error Loading "
    End If
    Erase filenames_list
    Exit Sub
hand_sav_mass:
    MsgBox "Error" + Str$(Err) + " saving " + filename_out, vbOKOnly, "Unknow error Saving"
    Erase filenames_list
End Sub

Private Sub OpenTextureButton_Click()
    On Error GoTo hand_op
    Dim pattern As String
    Dim tex As TEXTexture
    
    pattern = "Any Image file(*.bmp;*.jpg;*.gif;*.ico;*.rle;*.wmf;*.emf)|*.bmp;*.jpg;*.gif;*.ico;*.wmf;*.emf|TEX texture|*.tex"
    CommonDialog1.Filter = pattern
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer   'Allow multiple files to be selected
    CommonDialog1.ShowOpen 'Display the Open File Common Dialog

    If (CommonDialog1.fileName <> "") Then
        If (UCase(Right(CommonDialog1.fileName, 3)) = "TEX") Then
            ReadTEXTexture tex, CommonDialog1.fileName
            LoadBitmapFromTEXTexture tex
            hDC_Tex = tex.hdc
            hbmp_Tex = tex.hbmp
            Width_Tex = tex.width
            Height_Tex = tex.height
            TextureViewer.AutoSize = False
            TextureViewer.AutoRedraw = False
            TextureViewer.width = (tex.width + BorderWidth) * 15
            TextureViewer.height = (tex.height + BorderWidth) * 15
            TextureViewer.Picture = Nothing
            TexLoadedQ = True
            TransparentFlagCheck.Value = IIf(tex.ColorKeyFlag = 1, vbChecked, vbUnchecked)
        Else
            Set TextureViewer.Picture = LoadPicture(CommonDialog1.fileName)
            TextureViewer.AutoRedraw = True
            TextureViewer.AutoSize = True
            If (TexLoadedQ) Then
                DeleteDC hDC_Tex
                DeleteObject hbmp_Tex
                TexLoadedQ = False
            End If
            
        End If
    End If
    
    Image2TEX.width = IIf(TextureViewer.width < MinWidth, _
                        MinWidth, TextureViewer.width)
    Dim h As Long
    h = MinHeight + TextureViewer.Top + TextureViewer.height
    Image2TEX.height = h
    
    Exit Sub
hand_op:
    If Err <> 32755 Then
        MsgBox "Error" + Str$(Err), vbOKOnly, "Unknow error Loading"
    End If
End Sub

Private Sub SaveTextureButton_Click()
    On Error GoTo hand_sav
    Dim pattern As String
    Dim Texture As TEXTexture
    Dim Bmp As BMPTexture
    Dim hdc As Long
    Dim hbmp As Long
    
    If TexLoadedQ Then
        hdc = hDC_Tex
        hbmp = hbmp_Tex
    Else
        hdc = TextureViewer.hdc
        hbmp = TextureViewer.Picture.handle
    End If
    
    pattern = "TEX texture|*.tex|BMP image|*.bmp"
    CommonDialog1.Filter = pattern
    CommonDialog1.CancelError = True
    CommonDialog1.ShowSave 'Display the Save File Common Dialog
   
    If (CommonDialog1.fileName <> "") Then
        If (UCase(Right(CommonDialog1.fileName, 3)) = "BMP") Then
            GetBMPTextureFromBitmap Bmp, hdc, hbmp
            WriteBMPTexture Bmp, CommonDialog1.fileName
        Else
            GetTEXTextureFromBitmap Texture, hdc, hbmp
            Texture.ColorKeyFlag = IIf(TransparentFlagCheck.Value = vbChecked, 1, 0)
            WriteTEXTexture Texture, CommonDialog1.fileName
        End If
        MsgBox "Done", vbOKOnly, "Done"
    End If
    Exit Sub
hand_sav:
    If Err <> 32755 Then
        MsgBox "Error" + Str$(Err), vbOKOnly, "Unknow error Saving"
    End If
    
End Sub

Private Sub TextureViewer_Paint()
    BitBlt TextureViewer.hdc, 0, 0, Width_Tex, Height_Tex, hDC_Tex, 0, 0, vbSrcCopy
End Sub
