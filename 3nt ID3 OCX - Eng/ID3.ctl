VERSION 5.00
Begin VB.UserControl ID3 
   BackColor       =   &H00FAE8DA&
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   510
   ScaleWidth      =   1995
   ToolboxBitmap   =   "ID3.ctx":0000
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   90
      Left            =   975
      TabIndex        =   3
      Top             =   225
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "wav && mp3 info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   465
      Left            =   750
      TabIndex        =   1
      Top             =   0
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H007E511F&
      FillColor       =   &H00808080&
      Height          =   480
      Left            =   0
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID3 +"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   570
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "3nity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E3C1AA&
      Height          =   465
      Left            =   450
      TabIndex        =   2
      Top             =   0
      Width           =   1740
   End
End
Attribute VB_Name = "ID3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim TAG2 As clsTv2

Public TAGv1Exists As Boolean
Public TAGv1Genre As String
Public TAGv1Title As String
Public TAGv1Artist As String
Public TAGv1Album As String
Public TAGv1Created As String
Public TAGv1Comment As String
Public TAGv1GenreIndex As Integer

Public TAGv2Exists As Boolean
Public TAGv2Track As String
Public TAGv2Genre As String
Public TAGv2Title As String
Public TAGv2Artist As String
Public TAGv2Album As String
Public TAGv2Created As String
Public TAGv2Comment As String
Public TAGv2Composer As String
Public TAGv2OriginalArtist As String
Public TAGv2Copyright As String
Public TAGv2Url As String
Public TAGv2EncodedBy As String

Public WAVLenghtInSeconds As Long
Public WAVLenght As String
Public WAVBitrate As Long
Public WAVFrequency As Long
Public WAVFileSizeInBytes As Long
Public WAVFileSize As String
Public WAVChannels As Long
Public WAVBits As Long

Public Mp3FileSizeInBytes As Long
Public Mp3FileSize As String

Public Mp3FileLenghtInSeconds As String
Public Mp3FileLenght As String

Public Mp3Bitrate As String
Public Mp3Frequency As String
Public Mp3Channels As String
Public Mp3Copyright As String
Public Mp3CRC As String
Public Mp3Emphasis As String
Public Mp3Layer As String
Public Mp3MPEG As String
Public Mp3Original As String

Public Event sError(Error As Integer)

Dim locArtist As String * 30
Dim locTitle As String * 30
Dim locAlbum As String * 30
Dim locYear As String * 4
Dim locComment As String * 30
Dim locGenre As String * 1

Dim iFileName As String

Public Sub ReadID3(FileName As String)
iFileName = FileName
TAGv2Track = ""
TAGv2Genre = ""
TAGv2Title = ""
TAGv2Artist = ""
TAGv2Album = ""
TAGv2Created = ""
TAGv2Comment = ""
TAGv2Composer = ""
TAGv2OriginalArtist = ""
TAGv2Copyright = ""
TAGv2Url = ""
TAGv2EncodedBy = ""
    
TAGv1Title = ""
TAGv1Artist = ""
TAGv1Album = ""
TAGv1Created = ""
TAGv1Comment = ""
TAGv1Genre = ""
On Error Resume Next

Mp3FileSizeInBytes = FileLen(FileName)

If Mp3FileSizeInBytes < 1000 Then
    Mp3FileSize = Mp3FileSizeInBytes & " B"
ElseIf Mp3FileSizeInBytes >= 1000 And Mp3FileSizeInBytes < 1000000 Then
    Mp3FileSize = (Int(Mp3FileSizeInBytes / 1024 * 100)) / 100 & " kB"
Else
    Mp3FileSize = (Int(Mp3FileSizeInBytes / 1024 / 1024 * 100)) / 100 & " MB"
End If

Dim cc As Mp3

getMP3Info FileName, cc

Mp3FileLenghtInSeconds = cc.Lenght

Dim SS As Integer
Dim MM As Integer
Dim hh As Long


MM = cc.Lenght / 60
If Mp3FileLenghtInSeconds - MM * 60 < 0 Then MM = MM - 1

SS = Mp3FileLenghtInSeconds - MM * 60

If SS < 10 Then
    Mp3FileLenght = MM & ":0" & SS
Else
    Mp3FileLenght = MM & ":" & SS
End If

Mp3Bitrate = cc.Bitrate
Mp3Frequency = cc.Emphasis
Mp3Channels = cc.Channels
Mp3Copyright = cc.Copyright
Mp3CRC = cc.CRC
Mp3Emphasis = cc.Emphasis
Mp3Layer = cc.Layer
Mp3MPEG = cc.MPEG
Mp3Original = cc.Original

On Error GoTo NapakaID3v2
Set TAG2 = New clsTv2

TAG2.readtag FileName
If TAG2.hastag = True Then
    TAGv2Track = TAG2.getframevalue(etrack)
    TAGv2Genre = TAG2.getframevalue(6)
    TAGv2Title = TAG2.getframevalue(etitle)
    TAGv2Artist = TAG2.getframevalue(eartist)
    TAGv2Album = TAG2.getframevalue(ealbum)
    TAGv2Created = TAG2.getframevalue(eyear)
    TAGv2Comment = TAG2.getframevalue(ecomment)
    TAGv2Composer = TAG2.getframevalue(ecomposer)
    TAGv2OriginalArtist = TAG2.getframevalue(eorigartist)
    TAGv2Copyright = TAG2.getframevalue(ecopyright)
    TAGv2Url = TAG2.getframevalue(eurl)
    TAGv2EncodedBy = TAG2.getframevalue(eencodedby)
    TAGv2Exists = True
Else
NapakaID3v2:
    TAGv2Exists = False
End If

Set TAGv2 = Nothing
On Error Resume Next

Dim Buf As String * 128
Dim tmpStr As String
Dim i As Byte

MP3Datoteka = FileName

MP3Size = FileLen(MP3Datoteka)

Open MP3Datoteka For Binary As #1
  
Dim agenre() As String
Dim A As Integer
agenre = Split(sGenreMatrix)
TAGv1Genre = ""
Get #1, MP3Size - 127, Buf
      If Format(Left(Buf, 3), "<") <> "tag" Then
            TAGv1Exists = False
      Else
            TAGv1Title = Trim(Mid(Buf, 4, 30))
            TAGv1Artist = Trim(Mid(Buf, 34, 30))
            TAGv1Album = Trim(Mid(Buf, 64, 30))
            TAGv1Created = Trim(Mid(Buf, 94, 4))
            TAGv1Comment = Trim(Mid(Buf, 98, 30))
            For i = 0 To 148
              If i = Trim(Asc(Mid$(Buf, 128, 1))) Then Exit For
            Next i
        
            TAGv1GenreIndex = Trim(Asc(Mid$(Buf, 128, 1)))
            TAGv1Exists = True
      End If
Close #1

TAG1.FileName = FileName

Label3.Caption = TAGv1Title
TAGv1Title = Label3.Caption

Label3.Caption = TAGv1Artist
TAGv1Artist = Label3.Caption

End Sub

Public Sub RemoveID3Tagv2()

If iFileName <> "" Then
    Set TAG2 = New clsTv2
    TAG2.IError = 0
    
    TAG2.removetag iFileName

    If TAG2.IError > 0 Then RaiseEvent sError(TAG2.IError)
Else
    RaiseEvent sError(8001)
End If

End Sub

Public Sub SaveID3()

If iFileName <> "" Then
    Set TAG2 = New clsTv2
    
    TAG2.IError = 0

    TAG2.setframevalue etitle, TAGv2Title
    TAG2.setframevalue ealbum, TAGv2Album
    TAG2.setframevalue eartist, TAGv2Artist
    TAG2.setframevalue egenre, TAGv2Genre
    TAG2.setframevalue eyear, TAGv2Created
    TAG2.setframevalue eencodedby, TAGv2EncodedBy
    TAG2.setframevalue ecomment, TAGv2Comment
    TAG2.setframevalue etrack, TAGv2Track
    TAG2.setframevalue eorigartist, TAGv2OriginalArtist
    TAG2.setframevalue ecomposer, TAGv2Composer
    TAG2.setframevalue eurl, TAGv2Url
    TAG2.setframevalue ecopyright, TAGv2Copyright
    TAG2.writetag iFileName

    If TAG2.IError > 0 Then RaiseEvent sError(TAG2.IError)

    Dim STAG1 As mod1.ID3v1
    mod1.IError = 0
    STAG1.Title = TAGv1Title
    STAG1.Artist = TAGv1Artist
    STAG1.Album = TAGv1Album
    STAG1.sYear = TAGv1Created
    STAG1.Comments = TAGv1Comment
    STAG1.Genre = TAGv1Genre
  
    ShraniID3 iFileName, STAG1
    
    If mod1.IError > 0 And TAG2.IError <> mod1.IError Then
        RaiseEvent sError(mod1.IError)
    End If
Else
    RaiseEvent sError(8001)
    
End If

End Sub

Private Sub UserControl_Resize()
UserControl.Width = 2000
UserControl.Height = 480
Shape1.Width = 2000
Shape1.Height = 480

End Sub

Public Sub ReadWAV(FileName As String)
WAVInfo FileName

WAVLenghtInSeconds = wInfo.Lenght
WAVBitrate = wInfo.kbps
WAVFrequency = wInfo.Frequency
WAVFileSizeInBytes = FileLen(FileName)
WAVChannels = wInfo.Channels
WAVBits = wInfo.bits

If WAVFileSizeInBytes < 1000 Then
    WAVFileSize = WAVFileSizeInBytes & " B"
ElseIf FileSizeInBytes >= 1000 And WAVFileSizeInBytes < 1000000 Then
    WAVFileSize = (Int(WAVFileSizeInBytes / 1024 * 100)) / 100 & " kB"
Else
    WAVFileSize = (Int(WAVFileSizeInBytes / 1024 / 1024 * 100)) / 100 & " MB"
End If

Dim SS As Integer
Dim MM As Integer
Dim hh As Long


MM = wInfo.Lenght / 60
If wInfo.Lenght - MM * 60 < 0 Then MM = MM - 1

SS = wInfo.Lenght - MM * 60

If SS < 10 Then
    WAVLenght = MM & ":0" & SS
Else
    WAVLenght = MM & ":" & SS
End If

End Sub


