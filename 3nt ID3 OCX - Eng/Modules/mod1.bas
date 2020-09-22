Attribute VB_Name = "mod1"
Public IError As Integer

Public Type VBR
  Rate As String
  Lenght As String
End Type

Public Type Mp3
  Bitrate As String
  Channels As String
  Copyright As String
  CRC As String
  Emphasis As String
  Frequency As String
  Layer As String
  Lenght As String
  MPEG As String
  Original As String
  Size As String
End Type

Private MP3Dolzina As Long
Private MP3Datoteka As String

Public Type ID3v1
    Title As String * 30
    Artist As String * 30
    Album As String * 30
    sYear  As String * 4
    Comments As String * 30
    Genre As String * 1
End Type

Type WAV
    Channels As Integer
    Frequency As String
    bits As Integer
    kbps As Long
    FileSize As String
    Lenght As Long
End Type
    
Public wInfo As WAV

Public id3Info As ID3v1

'Can't remember the author of this code...It's not mine (the Genre Matrix)...
Public GenreArray() As String

Public Const sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Spacex|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"

Public Function ShraniID3(FileName As String, mp As ID3v1)
On Error GoTo er1

Dim TaG As String * 3
Open FileName For Binary As #1
Get #1, FileLen(FileName) - 127, TaG
If TaG = "TAG" Then
Put #1, FileLen(FileName) - 124, mp
Else
Put #1, FileLen(FileName) - 127, "TAG"
Close #1
Call ShraniID3(FileName, mp)
End If
Close #1
Exit Function

er1:

IError = vbError

End Function

Public Function WAVInfo(ImeDatoteke As String) As Boolean
Dim MM, SS As String
Dim Riff As String * 4
Dim f1X As String
Dim f1 As Byte

wInfo.FileSize = FileLen(ImeDatoteke)

Open ImeDatoteke For Binary As #1
    Get #1, 1, Riff
    Get #1, 23, wInfo.Channels
    Get #1, 35, wInfo.bits
    Get #1, 25, f1
Close #1

f1X = ConvertBase(val(Str(f1)), 10, 16)

If Riff <> "RIFF" Then WAVInfo = False: Exit Function

WAVInfo = True

Select Case f1X
    Case 40
    wInfo.Frequency = "8000"
    Case 11
    wInfo.Frequency = "11025"
    Case 22
    wInfo.Frequency = "22050"
    Case 0
    wInfo.Frequency = "32000"
    Case 44
    wInfo.Frequency = "44100"
    Case 80
    wInfo.Frequency = "48000"
End Select
    
wInfo.kbps = (CLng(wInfo.bits) * CLng(wInfo.Channels) * CLng(wInfo.Frequency))

wInfo.Lenght = (((wInfo.FileSize * 8) - 8000) / wInfo.kbps)


End Function

Public Function ConvertBase(NumIn As String, BaseIn As Integer, _
    BaseOut As Integer) As String
    
    Dim i As Integer, CurrentCharacter As String, _
    CharacterValue As Integer, PlaceValue As Integer, _
    RunningTotal As Double, Remainder As Double, _
    BaseOutDouble As Double, NumInCaps As String

    If NumIn = "" Or BaseIn < 2 Or BaseIn > 36 Or _
    BaseOut < 1 Or BaseOut > 36 Then
    ConvertBase = "Error"
    Exit Function
End If

NumInCaps = UCase$(NumIn)

PlaceValue = Len(NumInCaps)

For i = 1 To Len(NumInCaps)
    PlaceValue = PlaceValue - 1
    CurrentCharacter = Mid$(NumInCaps, i, 1)
    CharacterValue = 0
    If Asc(CurrentCharacter) > 64 And _
    Asc(CurrentCharacter) < 91 Then _
    CharacterValue = Asc(CurrentCharacter) - 55


    If CharacterValue = 0 Then
        If Asc(CurrentCharacter) < 48 Or _
        Asc(CurrentCharacter) > 57 Then
        ConvertBase = "Error"
        Exit Function
    Else
        CharacterValue = val(CurrentCharacter)
    End If
End If


If CharacterValue < 0 Or CharacterValue > BaseIn - 1 Then
    ConvertBase = "Error"
    Exit Function
End If
RunningTotal = RunningTotal + CharacterValue * _
(BaseIn ^ PlaceValue)
Next i

Do
BaseOutDouble = CDbl(BaseOut)
Remainder = ModDouble(RunningTotal, BaseOutDouble)
RunningTotal = (RunningTotal - Remainder) / BaseOut


If Remainder >= 10 Then
    CurrentCharacter = Chr$(Remainder + 55)
Else
    CurrentCharacter = Right$(Str$(Remainder), _
    Len(Str$(Remainder)) - 1)
End If
ConvertBase = CurrentCharacter & ConvertBase
Loop While RunningTotal > 0

End Function

Public Function ModDouble(NumIn As Double, DivNum As Double) As Double
ModDouble = NumIn - (Int(NumIn / DivNum) * DivNum)
    
End Function

Public Sub getMP3Info(ByVal lpMP3File As String, ByRef lpMP3Info As Mp3)
Dim Buf As String * 4096
Dim infoStr As String * 3
Dim lpVBRinfo As VBR
Dim tmpByte As Byte
Dim tmpNum As Byte
Dim i As Integer
Dim designator As Byte
Dim baseFreq As Single
Dim vbrBytes As Long

Open lpMP3File For Binary As #1
    Get #1, 1, Buf
Close #1

For i = 1 To 4092
    If Asc(Mid(Buf, i, 1)) = &HFF Then
        tmpByte = Asc(Mid(Buf, i + 1, 1))
        If Between(tmpByte, &HF2, &HF7) Or Between(tmpByte, &HFA, &HFF) Then
            Exit For
        End If
    End If
Next i

If i = 4093 Then

Else
      infoStr = Mid(Buf, i + 1, 3)
    
      tmpByte = Asc(Mid(infoStr, 1, 1))
    
      If ((tmpByte Mod 16) Mod 2) = 0 Then
            lpMP3Info.CRC = "Yes"
      Else
            lpMP3Info.CRC = "No"
      End If
    
      If Between(tmpByte, &HF2, &HF7) Then
            lpMP3Info.MPEG = "MPEG 2.0"
            designator = 1
      Else
            lpMP3Info.MPEG = "MPEG 1.0"
            designator = 2
      End If

    If Between(tmpByte, &HF2, &HF3) Or Between(tmpByte, &HFA, &HFB) Then
        lpMP3Info.Layer = "layer 3"
    Else
        If Between(tmpByte, &HF4, &HF5) Or Between(tmpByte, &HFC, &HFD) Then
            lpMP3Info.Layer = "layer 2"
        Else
            lpMP3Info.Layer = "layer 1"
        End If
    End If

    tmpByte = Asc(Mid(infoStr, 2, 1))
    
    If Between(tmpByte Mod 16, &H0, &H3) Then
        baseFreq = 22.05
    Else
        If Between(tmpByte Mod 16, &H4, &H7) Then
            baseFreq = 24
        Else
            baseFreq = 16
        End If
    End If
      
    lpMP3Info.Frequency = baseFreq * designator * 1000 & " Hz"

    tmpNum = tmpByte \ 16 Mod 16
    
    If designator = 1 Then
        If tmpNum < &H8 Then
            lpMP3Info.Bitrate = tmpNum * 8
        Else
            lpMP3Info.Bitrate = 64 + (tmpNum - 8) * 16
        End If
    Else
        If tmpNum <= &H5 Then
            lpMP3Info.Bitrate = (tmpNum + 3) * 8
        Else
            If tmpNum <= &H9 Then
                lpMP3Info.Bitrate = 64 + (tmpNum - 5) * 16
            Else
                If tmpNum <= &HD Then
                    lpMP3Info.Bitrate = 128 + (tmpNum - 9) * 32
                Else
                    lpMP3Info.Bitrate = 320
                End If
            End If
        End If
    End If
    
  MP3Dolzina = FileLen(lpMP3File) \ (val(lpMP3Info.Bitrate) / 8) \ 1000
  
    If Mid(Buf, i + 36, 4) = "Xing" Then
        vbrBytes = Asc(Mid(Buf, i + 45, 1)) * &H10000
        vbrBytes = vbrBytes + (Asc(Mid(Buf, i + 46, 1)) * &H100&)
        vbrBytes = vbrBytes + Asc(Mid(Buf, i + 47, 1))
        GetVBRrate lpMP3File, vbrBytes, lpVBRinfo
        lpMP3Info.Bitrate = lpVBRinfo.Rate
        lpMP3Info.Lenght = lpVBRinfo.Lenght
    Else
        lpMP3Info.Bitrate = lpMP3Info.Bitrate
        lpMP3Info.Lenght = MP3Dolzina
    End If
  
    tmpByte = Asc(Mid(infoStr, 3, 1))
    tmpNum = tmpByte Mod 16

    If tmpNum \ 8 = 1 Then
        lpMP3Info.Copyright = "Yes"
        tmpNum = tmpNum - 8
    Else
        lpMP3Info.Copyright = "No"
    End If

    If (tmpNum \ 4) Mod 2 Then
        lpMP3Info.Original = "Yes"
        tmpNum = tmpNum - 4
    Else
        lpMP3Info.Original = "No"
    End If

    Select Case tmpNum
        Case 0
            lpMP3Info.Emphasis = "None"
        Case 1
            lpMP3Info.Emphasis = "50/15 microseconds"
        Case 2
            lpMP3Info.Emphasis = "unvalid"
        Case 3
            lpMP3Info.Emphasis = "CITT j. 17"
    End Select

    tmpNum = (tmpByte \ 16) \ 4
    Select Case tmpNum
        Case 0
            lpMP3Info.Channels = "Stereo"
        Case 1
            lpMP3Info.Channels = "Joint Stereo"
        Case 2
            lpMP3Info.Channels = "2 Channel"
        Case 3
            lpMP3Info.Channels = "Mono"
    End Select
End If

lpMP3Info.Size = FileLen(lpMP3File)

End Sub

Private Sub GetVBRrate(ByVal lpMP3File As String, ByVal byteRead As Long, ByRef lpVBRinfo As VBR)
Dim i As Long
Dim OK As Boolean

i = 0
byteRead = byteRead - &H39
Do
    If byteRead > 0 Then
        i = i + 1
        byteRead = byteRead - 38 - Deljivo(i)
    Else
        OK = True
    End If
    
Loop Until OK

lpVBRinfo.Lenght = Trim(Str(i))
lpVBRinfo.Rate = Trim(Str(Int(8 * FileLen(lpMP3File) / (1000 * i))))

End Sub

Private Function Deljivo(ByVal Num As Long) As Byte
If Num Mod 3 = 0 Then
    Deljivo = 1
Else
    Deljivo = 0
End If
  
End Function

Public Function Between(ByVal accNum As Byte, ByVal accDown As Byte, ByVal accUp As Byte) As Boolean
If accNum >= accDown And accNum <= accUp Then
    Between = True
Else
    Between = False
End If

End Function
