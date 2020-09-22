VERSION 5.00
Object = "*\A3ntID3PlusOCX.vbp"
Begin VB.Form frmID3Info 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmID3Info"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin ID3PlusOCX.ID3 ID3 
      Left            =   7800
      Top             =   3600
      _ExtentX        =   3519
      _ExtentY        =   847
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3390
      Left            =   7275
      TabIndex        =   41
      Top             =   75
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   42
         Top             =   300
         Width           =   45
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Copy from Tag v1 to Tag v2"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7275
      TabIndex        =   40
      Top             =   5550
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy from Tag v2 to Tag v1"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7275
      TabIndex        =   39
      Top             =   5925
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7275
      TabIndex        =   38
      Top             =   6600
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "ID3 Tag v2"
      Height          =   4590
      Index           =   1
      Left            =   3075
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   4065
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   16
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   36
         Top             =   4125
         Width           =   2865
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   15
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   34
         Top             =   3750
         Width           =   2865
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   14
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   32
         Top             =   3375
         Width           =   2865
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   13
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   30
         Top             =   3000
         Width           =   2865
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   12
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   28
         Top             =   2625
         Width           =   2865
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   11
         Left            =   1050
         TabIndex        =   21
         Top             =   300
         Width           =   2865
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   10
         Left            =   1050
         TabIndex        =   20
         Top             =   675
         Width           =   2865
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   9
         Left            =   1050
         TabIndex        =   19
         Top             =   1050
         Width           =   2865
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   8
         Left            =   1050
         MaxLength       =   6
         TabIndex        =   18
         Top             =   1425
         Width           =   615
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   7
         Left            =   2400
         TabIndex        =   17
         Top             =   1425
         Width           =   1515
      End
      Begin VB.TextBox txt 
         Height          =   765
         Index           =   6
         Left            =   1050
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1800
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Encoded By:"
         Height          =   315
         Index           =   16
         Left            =   75
         TabIndex        =   37
         Top             =   4170
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "URL:"
         Height          =   315
         Index           =   15
         Left            =   75
         TabIndex        =   35
         Top             =   3795
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright:"
         Height          =   315
         Index           =   14
         Left            =   75
         TabIndex        =   33
         Top             =   3420
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Orig. Artist:"
         Height          =   315
         Index           =   13
         Left            =   75
         TabIndex        =   31
         Top             =   3045
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Composer:"
         Height          =   315
         Index           =   12
         Left            =   75
         TabIndex        =   29
         Top             =   2670
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Title:"
         Height          =   315
         Index           =   11
         Left            =   75
         TabIndex        =   27
         Top             =   350
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Artist:"
         Height          =   315
         Index           =   10
         Left            =   75
         TabIndex        =   26
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Album:"
         Height          =   315
         Index           =   9
         Left            =   75
         TabIndex        =   25
         Top             =   1095
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Year:"
         Height          =   315
         Index           =   8
         Left            =   75
         TabIndex        =   24
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Genre:"
         Height          =   315
         Index           =   7
         Left            =   1425
         TabIndex        =   23
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Comment:"
         Height          =   315
         Index           =   6
         Left            =   75
         TabIndex        =   22
         Top             =   1845
         Width           =   915
      End
   End
   Begin VB.FileListBox File1 
      Height          =   3600
      Left            =   150
      TabIndex        =   2
      Top             =   3375
      Width           =   2790
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   150
      TabIndex        =   1
      Top             =   525
      Width           =   2790
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   2790
   End
   Begin VB.Frame Frame1 
      Caption         =   "ID3 Tag v1"
      Height          =   2265
      Index           =   0
      Left            =   3075
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   4065
      Begin VB.ComboBox Genre 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmID3Info.frx":0000
         Left            =   2400
         List            =   "frmID3Info.frx":0282
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1425
         Width           =   1530
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   5
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   13
         Top             =   1800
         Width           =   2865
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   3
         Left            =   1050
         MaxLength       =   4
         TabIndex        =   10
         Top             =   1425
         Width           =   615
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   2
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1050
         Width           =   2865
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   1
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   6
         Top             =   675
         Width           =   2865
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   0
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   4
         Top             =   300
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Comment:"
         Height          =   315
         Index           =   5
         Left            =   75
         TabIndex        =   14
         Top             =   1845
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Genre:"
         Height          =   315
         Index           =   4
         Left            =   1425
         TabIndex        =   12
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Year:"
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   11
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Album:"
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   9
         Top             =   1095
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Artist:"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   7
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Title:"
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   5
         Top             =   350
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmID3Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

ID3.TAGv1Title = txt(0)
ID3.TAGv1Artist = txt(1)
ID3.TAGv1Album = txt(2)
ID3.TAGv1Created = txt(3)
If Genre.ListIndex = 0 Then
    ID3.TAGv1Genre = Chr(255)
Else
    ID3.TAGv1Genre = Chr(Genre.ItemData(Genre.ListIndex - 1))
End If
  
ID3.TAGv1Comment = txt(5)

ID3.TAGv2Comment = txt(6)
ID3.TAGv2Genre = txt(7)
ID3.TAGv2Created = txt(8)
ID3.TAGv2Album = txt(9)
ID3.TAGv2Artist = txt(10)
ID3.TAGv2Title = txt(11)
ID3.TAGv2Composer = txt(12)
ID3.TAGv2OriginalArtist = txt(13)
ID3.TAGv2Copyright = txt(14)
ID3.TAGv2Url = txt(15)
ID3.TAGv2EncodedBy = txt(16)

ID3.SaveID3

End Sub

Private Sub Command2_Click()
txt(0) = txt(11)
txt(1) = txt(10)
txt(2) = txt(9)
txt(3) = txt(8)
txt(5) = txt(6)

End Sub

Private Sub Command3_Click()
txt(11) = txt(0)
txt(10) = txt(1)
txt(9) = txt(2)
txt(8) = txt(3)
txt(7) = txt(4)
txt(6) = txt(5)

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_DblClick()
Dim iFileName As String

If Right(File1.Path, 1) <> "\" Then
    iFileName = File1.Path & "\" & File1.List(File1.ListIndex)
Else
    iFileName = File1.Path & File1.List(File1.ListIndex)
End If
    
If Right(File1.List(File1.ListIndex), 4) = ".mp3" Then
    Frame1(0).Visible = True
    Frame1(1).Visible = True
    
    Frame2.Left = Frame1(0).Left + Frame1(0).Width + 150
    Frame2.Caption = "MPEG Info"
    Frame2.Visible = True

    ID3.ReadID3 iFileName
    Genre.ListIndex = 0
    txt(0) = ID3.TAGv1Title
    txt(1) = ID3.TAGv1Artist
    txt(2) = ID3.TAGv1Album
    txt(3) = ID3.TAGv1Created
    Dim iCnt As Integer
    For iCnt = 0 To 148
        If ID3.TAGv1GenreIndex = Genre.ItemData(iCnt) Then
            Genre.ListIndex = iCnt + 1
        Else

        End If
    Next iCnt

    txt(5) = ID3.TAGv1Comment
    
    txt(6) = ID3.TAGv2Comment
    txt(7) = ID3.TAGv2Genre
    txt(8) = ID3.TAGv2Created
    txt(9) = ID3.TAGv2Album
    txt(10) = ID3.TAGv2Artist
    txt(11) = ID3.TAGv2Title
    txt(12) = ID3.TAGv2Composer
    txt(13) = ID3.TAGv2OriginalArtist
    txt(14) = ID3.TAGv2Copyright
    txt(15) = ID3.TAGv2Url
    txt(16) = ID3.TAGv2EncodedBy
    
    On Error Resume Next

    For iCnt = 0 To 10
        Load lbl(iCnt)
        lbl(iCnt).Top = (iCnt * 13 + 15) * Screen.TwipsPerPixelY
        lbl(iCnt).Left = lbl(0).Left
        lbl(iCnt).Visible = True
        lbl(iCnt).Caption = ""
    Next iCnt
    
    lbl(0).Caption = "Size: " & ID3.Mp3FileSize & "   (" & ID3.Mp3FileSizeInBytes & " B)"
    lbl(1).Caption = "Lenght: " & ID3.Mp3FileLenght & "   (" & ID3.Mp3FileLenghtInSeconds & " sec)"
    lbl(2).Caption = ID3.Mp3MPEG & " " & ID3.Mp3Layer
    lbl(3).Caption = ID3.Mp3Bitrate & " kbps"
    lbl(4).Caption = ID3.Mp3Frequency & " " & ID3.Mp3Channels
    lbl(5).Caption = "CRCs: " & ID3.Mp3CRC
    lbl(6).Caption = "Copyright: " & ID3.Mp3Copyright
    lbl(7).Caption = "Original: " & ID3.Mp3Original
    lbl(8).Caption = "Emphasis: " & ID3.Mp3Emphasis
    lbl(9).Caption = "ID3 Tag v1: " & ID3.TAGv1Exists
    lbl(10).Caption = "ID3 Tag v2: " & ID3.TAGv2Exists
    
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    
ElseIf Right(File1.List(File1.ListIndex), 4) = ".wav" Or Right(File1.List(File1.ListIndex), 4) = ".WAV" Then
    Frame1(0).Visible = False
    Frame1(1).Visible = False
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    
    Frame2.Left = Frame1(0).Left
    Frame2.Caption = "WAV Info"
    Frame2.Visible = True
    
    ID3.ReadWAV iFileName
    
    On Error Resume Next

    For iCnt = 0 To 10
        Load lbl(iCnt)
        lbl(iCnt).Top = (iCnt * 13 + 15) * Screen.TwipsPerPixelY
        lbl(iCnt).Left = lbl(0).Left
        lbl(iCnt).Visible = True
        lbl(iCnt).Caption = ""
    Next iCnt
    
    lbl(0).Caption = "Size: " & ID3.WAVFileSize & "   (" & ID3.WAVFileSizeInBytes & " B)"
    lbl(1).Caption = "Lenght: " & ID3.WAVLenght & "   (" & ID3.WAVLenghtInSeconds & " sec)"
    lbl(2).Caption = "Bitrate: " & ID3.WAVBitrate
    lbl(3).Caption = "Bits: " & ID3.WAVBits
    lbl(4).Caption = "Channels: " & ID3.WAVChannels
    lbl(5).Caption = "Frequency: " & ID3.WAVFrequency
    
Else
    MsgBox "Select *.mp3 or *.wav file!!!"
    
End If

End Sub

Private Sub ID3_sError(Error As Integer)
MsgBox "Error: " & Error

End Sub
