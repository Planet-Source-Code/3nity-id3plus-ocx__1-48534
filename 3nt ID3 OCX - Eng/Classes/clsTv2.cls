VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'I'm sorry to the autor of most of this code in this class...
'I tried to put the calss together as it was but when i was
'using it in my project (before creating an ocx) i deleted
'the coments:(

Option Explicit
Option Base 0

Public IError As Integer

Public Enum eid3frameids
    etrack = 1
    etitle = 2
    eartist = 3
    ealbum = 4
    eyear = 5
    egenre = 6
    ecomment = 7
    ecomposer = 8
    eorigartist = 9
    ecopyright = 10
    eurl = 11
    eencodedby = 12
End Enum

'enum lists all the supported frame types
Private Enum eframetypes
    etextframe = 1
    ewebframe = 2
    ecommentframe = 3
    eunknownframe = 4
End Enum

'10 byte structure to hold the id3v2.3 tag header info
Private Type tid3tagheader
    identifer As String * 3
    majorversion As Byte
    minorversion As Byte
    flags As Byte
    Size(3) As Byte
End Type

'10 byte structure to hold a frame header
Private Type tid3frameheader
    frameid As String * 4
    Size(3) As Byte
    flags(1) As Byte
End Type



'enum for the bit shifting direction
Private Enum ebitshiftdir
    eshiftleft = 1
    eshiftright = 2
End Enum

'tag identifiers

'__________s u p p o r t e d t a g s_________________
Private Const c_track = "TRCK"
Private Const c_title = "TIT2"
Private Const c_artist = "TPE1"
Private Const c_album = "TALB"
Private Const c_year = "TYER"
Private Const c_genre = "TCON"
Private Const c_comment = "COMM"
Private Const c_composer = "TCOM"
Private Const c_origartist = "TOPE"
Private Const c_copyright = "TCOP"
Private Const c_url = "WXXX"
Private Const c_encodedby = "TENC"

'__________u n s u p p o r t e d  t a g s____________
Private Const c_audioencryption = "AENC"
Private Const c_attachedpicture = "APIC"
Private Const c_commericalframe = "COMR"
Private Const c_encryptionmethodreg = "ENCR"
Private Const c_equalization = "EQUA"
Private Const c_eventtimecodes = "ETCO"
Private Const c_generalencapsulatedobj = "GEOB"
Private Const c_groupidentreg = "GRID"
Private Const c_involvedpeoplelist = "IPLS"
Private Const c_linkedinformation = "LINK"
Private Const c_musiccdidentifier = "MCDI"
Private Const c_mpeglocationlookuptable = "MLLT"
Private Const c_ownershipframe = "OWNE"
Private Const c_privateframe = "PRIV"
Private Const c_playcounter = "PCNT"
Private Const c_popularimeter = "POPM"
Private Const c_positionsyncframe = "POSS"
Private Const c_recommendedbuffersize = "RBUF"
Private Const c_relativevoladjust = "RVAD"
Private Const c_reverb = "RVRB"
Private Const c_synclyrictext = "SYLT"
Private Const c_synctempocodes = "SYTC"
Private Const c_beatsperminute = "TBPM"

'__________t e x t  t a g s__________________________
' these tags could be implemented quickly because they use the clsTF class
Private Const c_date = "TDAT"
Private Const c_playlistdelay = "TDLY"
Private Const c_lyricist = "TEXT"
Private Const c_filetype = "TFLT"
Private Const c_time = "TIME"
Private Const c_contentgroupdesc = "TIT1"
Private Const c_subtitle = "TIT3"
Private Const c_initalkey = "TKEY"
Private Const c_languages = "TLAN"
Private Const c_mediatype = "TMED"
Private Const c_origalbumtitle = "TOAL"
Private Const c_origfilename = "TOFN"
Private Const c_origlyricist = "TOLY"
Private Const c_origreleaseyear = "TORY"
Private Const c_fileowner = "TOWN"
Private Const c_bandorchestraaccomp = "TPE2"
Private Const c_conductor = "TPE3"
Private Const c_interpretedremixed = "TPE4"
Private Const c_partofset = "TPOS"
Private Const c_publisher = "TPUB"
Private Const c_recordingdates = "TRDA"
Private Const c_internetradiostationname = "TRSN"
Private Const c_internetradiostationowner = "TRSO"
Private Const c_size = "TSIZ"
Private Const c_internationstandardreccodes = "TSRC"
Private Const c_softwarehardwaresettings = "TSSE"
Private Const c_userdefinedtext = "TXXX"

'_____m o r e  u n s u p p o r t e d  t a g s________
Private Const c_uniquefileidentifer = "UFID"
Private Const c_termsofuser = "UFID"
Private Const c_unsynclyric = "USLT"

'_________________w e b  t a g s_____________________
' these tags could be implemented quickly because they use the clsWF class
Private Const c_commercialinfo = "WCOM"
Private Const c_copyrightlegalinfo = "WCOP"
Private Const c_officalaudiofilewebsite = "WOAF"
Private Const c_officalartistwebpage = "WOAR"
Private Const c_officalaudiosource = "WOAS"
Private Const c_payment = "WPAY"
Private Const c_publishersofficalwebpage = "WPUB"

'frames collection
Private m_objframes As clsF2

'private member variables
Private m_bhastag As Boolean

'-----------------------------------------------------------------------------------------------------
'   public properties
'-----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'   purpose: return a flag to indicate whether the mp3 has an id3v2 tag
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Public Property Get hastag() As Boolean
Attribute hastag.VB_Description = "Return a flag indicating whether or not the MP3 file contains an ID3v2.3.x Tag"
    hastag = m_bhastag
End Property


'-----------------------------------------------------------------------------------------------------
'   public methods and functions
'-----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'   purpose: parse the id3 tag for the file a_sfile
'   require: a_sfile <> ""
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Public Sub readtag(ByVal a_sfile As String)
Attribute readtag.VB_Description = "Read the contains for an MP3 file"

Dim fso As Scripting.FileSystemObject
Dim ifilenum As Integer

Dim tagheader As tid3tagheader
Dim ltagsize As Long

Dim ltagpos As Long
Dim tframeheader As tid3frameheader
Dim aframeheader(9) As Byte
Dim aframedata() As Byte
Dim lframesize As Long

Dim atag() As Byte
Dim ldx  As Long
Dim loffset As Long

On Error GoTo errhandler
    
    m_objframes.clear
    
    'first make sure the file exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(a_sfile) Then
        'call raiseerror(eerr_invalid_filename, "ID3v2::ReadTag", array("%FILE%"), array(a_sfile))
    End If
    
    ifilenum = FreeFile
    
    Open a_sfile For Binary As ifilenum
    
    Get ifilenum, 1, tagheader 'starting at the first byte of the file grab the tag header
    
    'check to see if there is an id3v2 tag present
    If tagheader.identifer <> "ID3" Then m_bhastag = False: GoTo normalexit Else m_bhastag = True
    
    'check to see what version the tag is.  we know how to read version id3v2.3.x
    If tagheader.majorversion <> 3 Then m_bhastag = False: GoTo normalexit
           
    'extract the tag size
    ltagsize = bytetolong(tagheader.Size)
    
    'grab the entire tag and close the file
    ReDim atag(ltagsize - 1)  'zero based array
    Get ifilenum, 11, atag
    Close ifilenum
    
    'we now have the entire tag, we now need to extract the frames from it
    ltagpos = 0
    Do While ltagpos < ltagsize
        
        'step 1: extract the 10 byte frame header
        'make sure we have enough bytes left to extract a tag header from
        
        If (ltagsize - ltagpos) < Len(tframeheader) Then Exit Do
        
        For ldx = ltagpos To ldx + Len(tframeheader) - 1
            aframeheader(ldx - ltagpos) = atag(ldx)
        Next ldx
        
        'move the tag position index forward according to the size of the frame header
        ltagpos = ltagpos + Len(tagheader)
        
        tframeheader = extractframeheader(aframeheader)
             
        'make sure we have not encountered padding
        If InStr(Trim$(tframeheader.frameid), vbNullChar) Then Exit Do
        
        'if not isvalidframeheader(tframeheader.frameid) then exit do
             
        'step 2: according to the size param in the frame header extract the frames data
                        
        lframesize = bytetolong(tframeheader.Size)
        
        loffset = ltagpos + lframesize - 1
        
        'size the data array according to the size of the frame data
        ReDim aframedata(lframesize - 1)
        
        'extract the frame data
        For ldx = ltagpos To loffset
            aframedata(ldx - ltagpos) = atag(ldx)
        Next ldx
        
        'we must increament the the tag position index according to the size of the tag data
        ltagpos = ltagpos + lframesize
        
        'step 3: process frame and add it to the frame collection
        
        'we need to determine what type of frame we are dealing with so we can
        Select Case lookupframetype(tframeheader.frameid)
        Case etextframe:
            Call extracttextframe(tframeheader, aframedata)
            
        Case ecommentframe:
            Call extractcommentframe(tframeheader, aframedata)
        
        Case ewebframe:
            Call extractwebframe(tframeheader, aframedata)
            
        Case eunknownframe:
            'skip over this frame...we do not know how to parse the frame data
            
        End Select
    Loop
        
normalexit:
On Error GoTo 0
Close ifilenum

Exit Sub

errhandler:

'make sure the file is closed!
Close ifilenum

'missatge ms_error, idi(156) & a_sfile & "'", idi(2), False
'err.raise err.number, err.source, err.description

            
End Sub

'----------------------------------------------------------------------------------------------------
'   purpose: write the id3 tag to file
'   require: a_sfile must exist
'   promise: nothing
'---
 '-------------------------------------------------------------------------------------------------
Public Sub writetag(ByVal a_sfile As String, Optional ByVal a_bignorereadonly As Boolean = True)
Attribute writetag.VB_Description = "Write the tag to an MP3 file"

    'this subroutine will extract the music data form the file a_sfile and store it in an array so
    'we can write all our changes to a temporary file before we replace a_sfile with the new tagged mp3.
    'by writing to the temp file we ensure that we do not damage the orignial mp3.
    'when the writing is successful we will replace a_file with the temp file.
    
Dim ifilenum

Dim lwritepos As Long

Dim lreadpos As Long
Dim lmusicdatasize As Long                  'this will hold the size of the music data the contanined in the mp3
Dim amusicdata() As Byte                    'we will copy the music data from the old mp3 to an array, this way we can write it to the end of our new mp3
Dim stmpfile As String                      'this holds the temporary file name we are going to preform all operations to.

Dim ttagheader As tid3tagheader
Dim ltagsize As Long                        'holds the tag size before the we preform the byte shifting
Dim atagsize(3) As Byte                     'this array will hold the 4 tag size bytes

Dim tframeheader As tid3frameheader
Dim aframeflags() As Byte

Dim lframesize As Long                      'holds the frame size before we preform the byte shifting
Dim aframesize(3) As Byte                   'this array will hold the 4 frame size bytes

Dim aframedata() As Byte                    'this array will contain the frames data

Dim lnumpaddingbytes As Long                'store the number of padding bytes to write at the end of the tag
Dim apadding() As Byte                      'ubound(apadding) = lnumpaddingbytes - 1 (contains the null chars)

Dim objframe As clsF

Dim fso As Scripting.FileSystemObject
Dim objfile As Scripting.File
Screen.MousePointer = 11
DoEvents
On Error GoTo errhandler
       
    
    'first check to make sure the file exists
    If Not doesfileexist(a_sfile) Then
        'call raiseerror(eerr_invalid_filename, "clsTv2::WriteTag", array("%FILE%"), array(a_sfile))
    End If
    
    'now check to see if the file is read only
    If isfilereadonly(a_sfile) Then
        
        'if the ignore flag is set we need to reset the read only flag
        If a_bignorereadonly Then
            Call SetAttr(a_sfile, vbNormal)
        Else
            'call raiseerror(eerr_fileisreadonly, "clsTv2::WriteTag", array("%FILE%"), array(a_sfile))
        End If
    End If
    
    'we made it here, it is safe to write the file
    
    '--------------------------------------------------------------------------------------------------------
    '   grab the music data from the file
    '--------------------------------------------------------------------------------------------------------
    ifilenum = FreeFile
    
    lreadpos = 1
    Open a_sfile For Binary As ifilenum
    
    Get ifilenum, lreadpos, ttagheader
    
    'check to see if there is a previous id3v2 tag
    If ttagheader.identifer = "ID3" Then
        'locate where the music data starts
        lreadpos = bytetolong(ttagheader.Size)
    Else
        'no tag, we can copy over the entire file
        lreadpos = 0
    End If
    
    lmusicdatasize = LOF(ifilenum) - lreadpos
    If lmusicdatasize < 0 Then
        'call raiseerror(eerr_invalid_filesize, "ID3v2::WriteTag", array("%FILE%"), array(a_sfile))
    End If
    
    'grab the music data
    ReDim amusicdata(lmusicdatasize - 1)  'zero based array
    
    Get ifilenum, lreadpos + 1, amusicdata
    
    Close ifilenum
    
    '--------------------------------------------------------------------------------------------------------
    '   write the id3v2 tag
    '--------------------------------------------------------------------------------------------------------
    
    ltagsize = 0
    lnumpaddingbytes = 0
    
    'step 1: build the tag header
    For Each objframe In m_objframes
        ltagsize = ltagsize + Len(tframeheader) + objframe.framesize
    Next objframe
    
    'we must make sure the tag size is divisable by 257 otherwise we will pad the mp3
    If ltagsize Mod 257 <> 0 Then
        
        lnumpaddingbytes = ltagsize Mod 257
        lnumpaddingbytes = 257 - lnumpaddingbytes
        
        ReDim apadding(lnumpaddingbytes - 1)    'the array is initalized to a byte value 0 or null
        
    End If
        
    With ttagheader
        .identifer = "ID3"
        .majorversion = 3
        .minorversion = 0
        
        Call longtobyte(ltagsize + lnumpaddingbytes, atagsize)
        Call copyarray(atagsize, .Size)
        
    End With
    
    
    'first we should generate a temp file to write the new tag to, then replace the old one
    stmpfile = gettempfile(a_sfile)
    
    'lets build the new tagged file
    ifilenum = FreeFile
    
    lwritepos = 1
    
    Open stmpfile For Binary As ifilenum
    
    Put ifilenum, lwritepos, ttagheader
    
    'advance the writing position
    lwritepos = lwritepos + Len(ttagheader)
    
    For Each objframe In m_objframes
        
        With tframeheader
            .frameid = objframe.frameid
                
            'copy over the frames flag data
            Call objframe.getflags(aframeflags)
            Call copyarray(aframeflags, .flags)
                            
            'calcuate the frame size
            lframesize = objframe.framesize
            Call longtobyte(lframesize, aframesize)
            Call copyarray(aframesize, .Size)
                
        End With
        
        Put ifilenum, lwritepos, tframeheader
        
        'advance the writing position
        lwritepos = lwritepos + Len(tframeheader)
        
        objframe.getbytes aframedata
        
        'write the tag data
        Put ifilenum, lwritepos, aframedata
        
        'advance the writing position
        lwritepos = lwritepos + UBound(aframedata) + 1  'zero based array
        
    Next objframe
    
    'check to see if there are any padding bytes to be written
    If lnumpaddingbytes > 0 Then
        Put ifilenum, lwritepos, apadding
    End If
    
    lwritepos = lwritepos + lnumpaddingbytes
    
    'finally write the audio data
    Put ifilenum, lwritepos, amusicdata
    
    Close ifilenum
    
    Set fso = New Scripting.FileSystemObject
    
    'delete the old mp3 file and rename the new one to the old file name
    Call fso.DeleteFile(a_sfile, True)
    Set objfile = fso.GetFile(stmpfile)
    
    objfile.Name = Mid$(a_sfile, InStrRev(a_sfile, "\") + 1)
        
    m_bhastag = True    'set the has tag flag
        
normalexit:
On Error GoTo 0
Set fso = Nothing
Screen.MousePointer = 0
Exit Sub

errhandler:
Screen.MousePointer = 0
IError = vbError
'clean up the temp file
If stmpfile <> "" Then Kill (stmpfile)

Set fso = Nothing

'pass the error up a level
Close ifilenum  'make sure the file is closes
'Err.Raise Err.Number, Err.Source, Err.description

End Sub

'----------------------------------------------------------------------------------------------------
'   purpose: remove the entire id3v2 tag from the mp3
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Public Sub removetag(ByVal a_sfile As String, Optional ByVal a_bignorereadonly As Boolean = True)
Attribute removetag.VB_Description = "Strip away an ID3v2 tag"
    
    'this subroutine will extract the music data from the a_sfile and copy it to an array.
    'it will write the music data to a temp file, delete a_sfile and rename the temp file
    'to a_sfile.  using the temp file ensures that we do not lose the mp3's data.
    
Dim ifilenum As Integer
Dim stmpfile As String                      'temp file name (this is the file all operations will be preformed on)

Dim ttagheader As tid3tagheader
Dim ltagsize As Long                        'the size of the tag we will be removing

Dim lmusicdatasize As Long                  'the size of the music data in the mp3
Dim amusicdata() As Byte                    'this array holds all the audio data ( ubound(amusicdata) = lmusicdatasize -1

Dim fso As Scripting.FileSystemObject
Dim objfile As Scripting.File

On Error GoTo errhandler

'first check to make sure the file exists
    If Not doesfileexist(a_sfile) Then
        'call raiseerror(eerr_invalid_filename, "clsTv2::RemoveTag", array("%FILE%"), array(a_sfile))
    End If
    
    'now check to see if the file is read only
    If isfilereadonly(a_sfile) Then
        
        'if the ignore flag is set we need to reset the read only flag
        If a_bignorereadonly Then
            Call SetAttr(a_sfile, vbNormal)
        Else
            'call raiseerror(eerr_fileisreadonly, "clsTv2::WriteTag", array("%FILE%"), array(a_sfile))
        End If
    End If
    
    ifilenum = FreeFile
    
    Open a_sfile For Binary As ifilenum
    
    Get ifilenum, 1, ttagheader
    
    'check to see if there is a tag to remove
    If ttagheader.identifer <> "ID3" Then Exit Sub
    
    ltagsize = bytetolong(ttagheader.Size)
    
    lmusicdatasize = LOF(ifilenum) - ltagsize
    
    'check to make sure the music data size is valid
    If lmusicdatasize > LOF(ifilenum) Then
        'call raiseerror(eerr_invalid_filesize, "clsTv2::RemoveTag", array("%FILE%"), array(a_sfile))
    End If
    
    ReDim amusicdata(lmusicdatasize - 1) 'zero based array
    
    Get ifilenum, ltagsize + 1, amusicdata
    
    Close ifilenum
        
    'copy the data to a temp file first
    stmpfile = gettempfile(a_sfile)
    
    ifilenum = FreeFile
    
    'write the temp file
    Open stmpfile For Binary As ifilenum
    Put ifilenum, 1, amusicdata
    Close ifilenum
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'delete the old version of the mp3
    Call fso.DeleteFile(a_sfile, True)
    
    'rename the new version
    Set objfile = fso.GetFile(stmpfile)
    
    objfile.Name = Mid$(a_sfile, InStrRev(a_sfile, "\") + 1)
    
    'clear out the frames collection and we are done!
    m_objframes.clear
    
    m_bhastag = False
    
normalexit:
On Error GoTo 0
Close ifilenum
Set fso = Nothing

Exit Sub

errhandler:
Close ifilenum
Me.IError = vbError

'clean up the temp file
If stmpfile <> "" Then Kill (stmpfile)

Set fso = Nothing

End Sub

'----------------------------------------------------------------------------------------------------
'   purpose: look through the frames collection for the frame with a frame id matching eframeid
'            and return it's value
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Public Function getframevalue(ByVal a_eframeid As eid3frameids) As String
Attribute getframevalue.VB_Description = "Return a specified frames data value"
Dim objframe As clsF
Dim sframeid As String

    sframeid = lookupframeid(a_eframeid)
    
    For Each objframe In m_objframes
        If objframe.frameid = sframeid Then
            
            'remove any null char's
            getframevalue = Trim(Replace$(objframe.data, vbNullChar, " ")) 'afegit trim per ke no hi hagui grups i discos dobles
            Exit For
        End If
    Next objframe

End Function

'----------------------------------------------------------------------------------------------------
'   purpose: set a frames text value
'   require: nothing
'   promise: if the frame is not found one is created
'----------------------------------------------------------------------------------------------------
Public Sub setframevalue(ByVal a_eframeid As eid3frameids, ByVal a_sdata As String)
Attribute setframevalue.VB_Description = "Set a Frames text data"
Dim objframe As clsF
Dim sframeid As String

Dim btagfound As Boolean    'flag to track if the frame was found

    btagfound = False
    
    sframeid = lookupframeid(a_eframeid)
        
    For Each objframe In m_objframes
        If objframe.frameid = sframeid Then
            
            btagfound = True
            
            'set the new value
            objframe.data = a_sdata
            
            Exit For    'get out
        End If
    Next objframe
    
    'check to see if the tag was updated
    If Not btagfound Then
        
        Select Case lookupframetype(sframeid)
        Case etextframe
            Call addtextframe(sframeid, a_sdata)
            
        Case ewebframe
            Call addwebframe(sframeid, a_sdata)
            
        Case ecommentframe
            Call addcommentframe(sframeid, a_sdata)
            
        Case eunknownframe
            'call raiseerror(eerr_unknownframeid, "clsTv2::setFrameValue", array("%FRAMEID%"), array(sframeid))
            
        End Select
    End If
    
End Sub

'-----------------------------------------------------------------------------------------------------
'   private methods and functions
'-----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'   purpose: lookup the frame id
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Function lookupframeid(ByVal a_eid3frame As eid3frameids)

    Select Case a_eid3frame
    Case etrack:
        lookupframeid = c_track
        
    Case etitle:
        lookupframeid = c_title
        
    Case eartist:
        lookupframeid = c_artist
        
    Case ealbum:
        lookupframeid = c_album
        
    Case eyear:
        lookupframeid = c_year
            
    Case egenre:
        lookupframeid = c_genre
        
    Case ecomment:
        lookupframeid = c_comment
        
    Case ecomposer:
        lookupframeid = c_composer
        
    Case eorigartist:
        lookupframeid = c_origartist
        
    Case ecopyright:
        lookupframeid = c_copyright
        
    Case eurl:
        lookupframeid = c_url
        
    Case eencodedby:
        lookupframeid = c_encodedby
        
    Case Else
        'call raiseerror(eerr_unknownframeid, "ID3v2::LookupFrameID")
    End Select

End Function

'----------------------------------------------------------------------------------------------------
'   purpose: lookup the type of frame based on the frame id
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Function lookupframetype(ByVal sframeid As String) As eframetypes

    Select Case sframeid
    
    Case c_track, c_title, c_artist, c_album, c_year, c_genre, c_composer, c_origartist, c_copyright, c_encodedby
        lookupframetype = etextframe
        
    Case c_comment:
        lookupframetype = ecommentframe
        
    Case c_url:
        lookupframetype = ewebframe
        
    Case Else
        lookupframetype = eunknownframe
        
    End Select
    
End Function



'----------------------------------------------------------------------------------------------------
'   purpose: extract the frame header bytes to a tid3frameheader type
'   require: ubound(a_bytes) = c_cxframeheadersize
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Function extractframeheader(ByRef a_bytes() As Byte) As tid3frameheader
Dim idx As Integer
Dim sframeid As String

    If UBound(a_bytes) <> 9 Then
        'call raiseerror(eerr_invalid_frameheadersize, "ID3v2::ExtractFrameHeader")
    End If
    
    With extractframeheader
        
        'extract the frame id
        For idx = 0 To 3
            sframeid = sframeid & Chr$(a_bytes(idx))
        Next idx
        
        .frameid = sframeid
        
        'extract the frame size
        For idx = 4 To 7
            .Size(idx - 4) = a_bytes(idx)
        Next idx
        
        'extract the frame flags
        For idx = 8 To 9
            .flags(idx - 8) = a_bytes(idx)
        Next idx
    End With
    
    
End Function

'----------------------------------------------------------------------------------------------------
'   purpose: extract a text frame and add it to the frames collection
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Sub extracttextframe(ByRef theader As tid3frameheader, ByRef a_data() As Byte)
    
Dim objframe As clsF
    
    Set objframe = New clsTF
    
    With objframe
        .frameid = theader.frameid
        .setflags theader.flags
        
        .parsebytes a_data
    End With
    
    m_objframes.add objframe
    
End Sub

'----------------------------------------------------------------------------------------------------
'   purpose: extract a comment frame and add it to the frames collection
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Sub extractcommentframe(ByRef theader As tid3frameheader, ByRef a_data() As Byte)

Dim objframe As clsF

    Set objframe = New clsCF
    
    With objframe
        .frameid = theader.frameid
        .setflags theader.flags
        
        .parsebytes a_data
    End With
    
    m_objframes.add objframe

End Sub

'----------------------------------------------------------------------------------------------------
'   purpose: extract a new web frame and add it to the frames collection
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Sub extractwebframe(ByRef theader As tid3frameheader, ByRef a_data() As Byte)

Dim objframe As clsF

    Set objframe = New clsWF
    
    With objframe
        .frameid = theader.frameid
        .setflags theader.flags
        
        .parsebytes a_data
    End With

    m_objframes.add objframe

End Sub

'----------------------------------------------------------------------------------------------------
'   purpose: create a new text frame and add it to the collection
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Sub addtextframe(ByVal a_sframeid, ByVal a_sdata As String)
Dim objframe As clsF

    Set objframe = New clsTF
    
    With objframe
        .frameid = a_sframeid
        .data = a_sdata
    End With
    
    'add the frame to the collection
    Call m_objframes.add(objframe)
    
End Sub

'----------------------------------------------------------------------------------------------------
'   purpose: create a new comment frame and add it to the collection
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Sub addcommentframe(ByVal a_sframeid, ByVal a_sdata As String)

'currently there is no special code for comment frames, this method is here for future purposes
' i.e. allowing the user to set the language, or the description field

Dim objframe As clsF

    Set objframe = New clsCF
    
    With objframe
        .frameid = a_sframeid
        .data = a_sdata
    End With
    
    'add the frame to the collection
    Call m_objframes.add(objframe)
    
End Sub

'----------------------------------------------------------------------------------------------------
'   purpose: create a new web frame and add it to the collection
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Sub addwebframe(ByVal a_sframeid, ByVal a_sdata As String)

'currently there is no special code for web frames, this method is here for future purposes
' i.e. allowing the user to set the description field
Dim objframe As clsF

    Set objframe = New clsWF
    With objframe
        .frameid = a_sframeid
        .data = a_sdata
    End With
    
    'add the frame to the collection
    Call m_objframes.add(objframe)
    
End Sub

'----------------------------------------------------------------------------------------------------
'   purpose: copy the contents of one array to another
'   require: ubound(a_source) = ubound(a_dest)
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Sub copyarray(ByRef a_source() As Byte, ByRef a_dest() As Byte)
    
Dim idx As Integer

    If UBound(a_source) <> UBound(a_dest) Then
        'call raiseerror(eerr_arrayboundsdonotmatch, "clsTv2::CopyArray")
    End If
    
    For idx = LBound(a_source) To UBound(a_source)
        a_dest(idx) = a_source(idx)
    Next idx
    
    
End Sub


'----------------------------------------------------------------------------------------------------
'   initalization / deconstruction
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'   purpose: initialize the frame colllection
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Sub class_initialize()
    
    Set m_objframes = New clsF2
    m_bhastag = False
   
End Sub

'----------------------------------------------------------------------------------------------------
'   purpose: distroy the frame collection
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Sub class_terminate()
    Set m_objframes = Nothing
End Sub




'/---------------------------------------------------------------------------------------------------\
' file name:    id3v2enums.bas
' description:  enums and constants
'
' author:       r. glenn scott
' e-mail:       glenn@ingeo.org
'
' language:     visual basic 6.0 sp5
' written on:   july 9, 2001
'
' copyright (c) r. glenn scott
' the copyright to the computer program(s) herein
' is the property of r. glenn scott, canada.
' the program(s) may be used and/or copied only with
' the written permission of r. glenn scott
' or in accordance with the terms and conditions
' stipulated in the agreement/contract under which
' the program(s) have been supplied.
'
' revision history
' =================
'   07/31/2001  -   fixed the tag size caculation methods (bytetolong & longtobyte)
'                   the algorithm implemented previously did not work with all tag sizes.
'
'   07/31/2001  -   added a bitshifting method (shiftbits)
'\---------------------------------------------------------------------------------------------------/


'----------------------------------------------------------------------------------------------------
'   purpose: id3v2 standard states that the 8th bit of every byte should be shifted for proper storage
'            of the frame size.  this function converts the frame size to a byte array with the
'            encoded byte size.
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Sub longtobyte(ByVal val As Long, ByRef bytearray() As Byte)

'gs07312001 -   replaced with the correct implementation
'dim byte1 as byte
'dim byte2 as byte
'dim byte3 as byte
'dim byte4 as byte
'
'    byte1 = val and 127
'    val = val / 128
'
'    byte2 = val and 127
'    val = val / 128
'
'    byte3 = val and 127
'    val = val / 128
'
'    byte4 = val and 127
'
'    redim bytearray(3)
'    bytearray(0) = byte4
'    bytearray(1) = byte3
'    bytearray(2) = byte2
'    bytearray(3) = byte1

Dim idx As Integer

On Error GoTo errhandler

    For idx = 0 To 3
        bytearray(idx) = shiftbits(val, (3 - idx) * 7, eshiftright) And 127
    Next idx

normalexit:
On Error GoTo 0
Exit Sub

errhandler:
'raise the error back to the caller
Err.Raise Err.Number, "ID3v2Enums::LongToByte", Err.description
    
End Sub

'----------------------------------------------------------------------------------------------------
'   purpose: extract the frame size back to a long
'   require: ubound(bytearray) = 3
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Function bytetolong(ByRef bytearray() As Byte) As Long

'gs07312001 -   replaced with a loop
'    bytetolong = bytearray(0) * (2 ^ 21)
'    bytetolong = bytetolong + bytearray(1) * (2 ^ 14)
'    bytetolong = bytetolong + bytearray(2) * (2 ^ 7)
'    bytetolong = bytetolong + bytearray(3) * (2 ^ 0)

Dim idx As Integer

On Error GoTo errhandler

    bytetolong = 0
    For idx = 0 To 3
        bytetolong = bytetolong + (bytearray(idx) * (2 ^ ((3 - idx) * 7)))
    Next idx

normalexit:
On Error GoTo 0
Exit Function

errhandler:
'raise the error back to the caller
Err.Raise Err.Number, "ID3v2Enums::ByteToLong", Err.description

End Function

'----------------------------------------------------------------------------------------------------
'   purpose: shift the orignal value by the number of bits specified
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Function shiftbits(ByVal lvalue As Long, ByVal lnumbitstoshift As Long, ByVal edir As ebitshiftdir) As Long

 
On Error GoTo errhandler
   
    Select Case edir
    
    Case eshiftleft: shiftbits = lvalue * (2 ^ lnumbitstoshift)
    
    Case eshiftright: shiftbits = lvalue \ (2 ^ lnumbitstoshift)
    
    End Select
        
normalexit:
On Error GoTo 0
Exit Function
    
errhandler:
Err.Raise Err.Number, "Shift", Err.description
Resume normalexit

End Function

'----------------------------------------------------------------------------------------------------
'   purpose: this function is used to determine if the 4 characters passed in make
'            a vaild id3v2.3 frame header.
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Function isvalidframeheader(ByVal a_sframeid As String) As Boolean
    
    isvalidframeheader = True
    
    Select Case a_sframeid
    '____________s u p p o r t e d  t a g s______________
    Case c_track
    Case c_title
    Case c_artist
    Case c_album
    Case c_year
    Case c_genre
    Case c_comment
    Case c_composer
    Case c_origartist
    Case c_copyright
    Case c_url
    Case c_encodedby
    
    '__________u n s u p p o r t e d  t a g s____________
    Case c_audioencryption = "AENC"
    Case c_attachedpicture = "APIC"
    Case c_commericalframe = "COMR"
    Case c_encryptionmethodreg = "ENCR"
    Case c_equalization = "EQUA"
    Case c_eventtimecodes = "ETCO"
    Case c_generalencapsulatedobj = "GEOB"
    Case c_groupidentreg = "GRID"
    Case c_involvedpeoplelist = "IPLS"
    Case c_linkedinformation = "LINK"
    Case c_musiccdidentifier = "MCDI"
    Case c_mpeglocationlookuptable = "MLLT"
    Case c_ownershipframe = "OWNE"
    Case c_privateframe = "PRIV"
    Case c_playcounter = "PCNT"
    Case c_popularimeter = "POPM"
    Case c_positionsyncframe = "POSS"
    Case c_recommendedbuffersize = "RBUF"
    Case c_relativevoladjust = "RVAD"
    Case c_reverb = "RVRB"
    Case c_synclyrictext = "SYLT"
    Case c_synctempocodes = "SYTC"
    Case c_beatsperminute = "TBPM"
    
    '__________t e x t  t a g s__________________________
    ' these tags could be implemented quickly because they use the clsTF class
    Case c_date = "TDAT"
    Case c_playlistdelay = "TDLY"
    Case c_lyricist = "TEXT"
    Case c_filetype = "TFLT"
    Case c_time = "TIME"
    Case c_contentgroupdesc = "TIT1"
    Case c_subtitle = "TIT3"
    Case c_initalkey = "TKEY"
    Case c_languages = "TLAN"
    Case c_mediatype = "TMED"
    Case c_origalbumtitle = "TOAL"
    Case c_origfilename = "TOFN"
    Case c_origlyricist = "TOLY"
    Case c_origreleaseyear = "TORY"
    Case c_fileowner = "TOWN"
    Case c_bandorchestraaccomp = "TPE2"
    Case c_conductor = "TPE3"
    Case c_interpretedremixed = "TPE4"
    Case c_partofset = "TPOS"
    Case c_publisher = "TPUB"
    Case c_recordingdates = "TRDA"
    Case c_internetradiostationname = "TRSN"
    Case c_internetradiostationowner = "TRSO"
    Case c_size = "TSIZ"
    Case c_internationstandardreccodes = "TSRC"
    Case c_softwarehardwaresettings = "TSSE"
    Case c_userdefinedtext = "TXXX"
    
    '_____m o r e  u n s u p p o r t e d  t a g s________
    Case c_uniquefileidentifer = "UFID"
    Case c_termsofuser = "UFID"
    Case c_unsynclyric = "USLT"
    
    '_________________w e b  t a g s_____________________
    ' these tags could be implemented quickly because they use the clsWF class
    Case c_commercialinfo = "WCOM"
    Case c_copyrightlegalinfo = "WCOP"
    Case c_officalaudiofilewebsite = "WOAF"
    Case c_officalartistwebpage = "WOAR"
    Case c_officalaudiosource = "WOAS"
    Case c_payment = "WPAY"
    Case c_publishersofficalwebpage = "WPUB"
    
    'if we make it here we have a non supported frame
    Case Else
    
        'frame id's starting with x, y, or z are experimental tags
        Select Case Left$(a_sframeid, 1)
        Case "X":
        Case "Y":
        Case "Z":
        Case Else:
            isvalidframeheader = False
        End Select
    
    End Select

End Function

'----------------------------------------------------------------------------------------------------
'   purpose: check to see if a file is flagged as read-only
'   require: file path to check
'   promise: returns true if the a_sfile is marked as read-only
'            if the file cannot be found an error is raised
'----------------------------------------------------------------------------------------------------
Private Function isfilereadonly(ByVal a_sfile As String) As Boolean
    If doesfileexist(a_sfile) Then
        If (GetAttr(a_sfile) And vbReadOnly) Then isfilereadonly = True Else isfilereadonly = False
    Else
        'call raiseerror(eerr_invalid_filename, "IsFileReadOnly", array("%FILE%"), array(a_sfile))
    End If
End Function

'----------------------------------------------------------------------------------------------------
'   purpose: check to see if a file is flagged as archived
'   require: file path to check
'   promise: returns true if the a_sfile is marked as archive
'            if the file cannot be found an error is raised
'----------------------------------------------------------------------------------------------------
Private Function isfilearchived(ByVal a_sfile As String) As Boolean
    If doesfileexist(a_sfile) Then
        If (GetAttr(a_sfile) And vbArchive) Then isfilearchived = True Else isfilearchived = False
    Else
        'call raiseerror(eerr_invalid_filename, "IsFileArchived", array("%FILE%"), array(a_sfile))
    End If
End Function

'----------------------------------------------------------------------------------------------------
'   purpose: check to see if a specified file exists
'   require: file path to check
'   promise: returns true is the file can be found
'----------------------------------------------------------------------------------------------------
Private Function doesfileexist(ByVal a_sfile As String) As Boolean
Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    If fso.FileExists(a_sfile) Then doesfileexist = True Else doesfileexist = False
    Set fso = Nothing
End Function


'----------------------------------------------------------------------------------------------------
'   purpose: generates a unique file
'   require: nothing
'   promise: nothing
'----------------------------------------------------------------------------------------------------
Private Function gettempfile(ByVal a_sfilename) As String


Dim fso As Scripting.FileSystemObject
Dim stmpfile As String

Dim sdir As String
Dim sfile As String

Dim sext As String

On Error GoTo errhandler
    
    Set fso = New Scripting.FileSystemObject
    
    sdir = Left$(a_sfilename, InStrRev(a_sfilename, "\"))
    sfile = Mid$(a_sfilename, InStrRev(a_sfilename, "\") + 1)
    sext = ".new"
    stmpfile = sdir & sfile & sext
    
    Do While fso.FileExists(stmpfile)
        stmpfile = sdir & Format(Now, "HhMmSs-") & sfile
    Loop
        
    'return the temp file name
    gettempfile = stmpfile
        
normalexit:
On Error GoTo 0
Set fso = Nothing

Exit Function

errhandler:
Set fso = Nothing

Err.Raise Err.Number, Err.Source, Err.description

End Function

