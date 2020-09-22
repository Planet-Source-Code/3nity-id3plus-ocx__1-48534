VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private m_sframeid As String
Private m_flags() As Byte

Public Property Let frameid(ByVal a_sframeid As String)
Attribute frameid.VB_Description = "Return/Set the Frame ID"
End Property

Public Property Get frameid() As String
End Property

Public Property Get framesize() As Long
Attribute framesize.VB_Description = "Return the Frame Size"
End Property

Public Sub setflags(ByRef a_flags() As Byte)
Attribute setflags.VB_Description = "Set the Frame's flags"
End Sub

Public Sub getflags(ByRef a_flags() As Byte)
Attribute getflags.VB_Description = "Return the frame's flags"
End Sub

Public Sub parsebytes(ByRef a_bytes() As Byte)
Attribute parsebytes.VB_Description = "Parse the frame's content"
End Sub

Public Sub getbytes(ByRef a_bytes() As Byte)
Attribute getbytes.VB_Description = "Return the frame contents in a byte array. (read to be written to file)"
End Sub

Public Property Get textencoding() As Byte
Attribute textencoding.VB_Description = "Set the Frame's text encoding field"
End Property

Public Property Let textencoding(ByVal a_enc As Byte)
End Property

Public Property Get data() As String
Attribute data.VB_Description = "Return/Set the data field"
End Property

Public Property Let data(ByVal a_sdata As String)
End Property
