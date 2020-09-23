VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   3840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4200
      Width           =   6615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Zip File Comment"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open ZIP File"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileEnd 
         Caption         =   "E&xit Program"
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private I As Long, Pos As Long
Private TempLong As Long, TempString As String
Private TempByte As Byte, TempByte2 As Byte
Private TempByte3 As Byte, TempByte4 As Byte

Private FileCount As Integer
Private LocalHeader() As LocalFileHeader
Private CentralHeader() As FileHeader
Private CurrentFile As String
Private EndCentral As EndOfCentralDirectory

Private Sub Form_Load()
Caption = App.Title
With ListView1
  .ColumnHeaders.Add , , "File Name"
  .ColumnHeaders.Add , , "Compression Method"
  .ColumnHeaders.Add , , "CRC-32"
  .ColumnHeaders.Add , , "Compressed Size"
  .ColumnHeaders.Add , , "Decompressed Size"
End With
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
'MsgBox (Item.Index - 1) & vbCrLf & vbCrLf & Hex(LocalHeader(Item.Index - 1).LocalFileHeaderSignature)
End Sub

Private Sub mnuFileEnd_Click()
Unload Me: End
End Sub

Private Sub mnuFileOpen_Click()
On Error Resume Next
ListView1.ListItems.Clear
FileCount = 0

CommonDialog1.DialogTitle = "Open ZIP File"
CommonDialog1.FileName = ""
CommonDialog1.Filter = "ZIP Format Compressed Files|*.zip"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub

CurrentFile = CommonDialog1.FileName
Open CommonDialog1.FileName For Binary Access Read As 1

  ProgressBar1.Max = LOF(1)
  MousePointer = vbHourglass
  For I = 1 To LOF(1)
    Get 1, I, TempByte
    If TempByte = 80 Then
      Get 1, I + 1, TempByte2
      If TempByte2 = 75 Then
        Get 1, I + 2, TempByte3
        If TempByte3 = 5 Then
          Get 1, I + 3, TempByte4
          If TempByte4 = 6 Then
            Get 1, I, EndCentral
            Exit For
          End If
        End If
      End If
    End If
  ProgressBar1.Value = I
  DoEvents
  Next I
            
  Text1.Text = Input(EndCentral.ZIPFileCommentLength, 1)
  ReDim CentralHeader(EndCentral.TotalNumberOfEntries - 1) As FileHeader
  ReDim LocalHeader(EndCentral.TotalNumberOfEntries - 1) As LocalFileHeader
  
  For I = 1 To LOF(1)
    Get 1, I, TempByte
    If TempByte = 80 Then
      Get 1, I + 1, TempByte2
      If TempByte2 = 75 Then
        Get 1, I + 2, TempByte3
        If TempByte3 = 1 Then
          Get 1, I + 3, TempByte4
          If TempByte4 = 2 Then
          
            Get 1, I, CentralHeader(FileCount)
            TempString = Input(CentralHeader(FileCount).FileNameLength, 1)
            If CentralHeader(FileCount).UncompressedSize <> 0 Then
              FileCount = FileCount + 1
              With ListView1
                .ListItems.Add FileCount, , TempString
                .ListItems(FileCount).SubItems(1) = ResolveMethod(CentralHeader(FileCount - 1).CompressionMethod)
                .ListItems(FileCount).SubItems(2) = Hex(CentralHeader(FileCount - 1).CRC32)
                .ListItems(FileCount).SubItems(3) = CentralHeader(FileCount - 1).CompressedSize & " bytes"
                .ListItems(FileCount).SubItems(4) = CentralHeader(FileCount - 1).UncompressedSize & " bytes"
              End With
            Else
              FileCount = FileCount + 1
              With ListView1
                .ListItems.Add FileCount, , TempString
                .ListItems(FileCount).SubItems(1) = ResolveMethod(CentralHeader(FileCount - 1).CompressionMethod)
                .ListItems(FileCount).SubItems(2) = "Directory Entry"
                .ListItems(FileCount).SubItems(3) = "Directory Entry"
                .ListItems(FileCount).SubItems(4) = "Directory Entry"
              End With
            End If
                        
          End If
        End If
      End If
    End If
    ProgressBar1.Value = I
    DoEvents
  Next I
  
  For I = 0 To UBound(CentralHeader())
    Seek 1, CentralHeader(I).RelativeOffsetOfLocalHeader + 1
    Get 1, , LocalHeader(I)
  Next I
  MousePointer = vbDefault
Close 1
End Sub
