VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements clsF

Private m_sframeid As String * 4
Private m_aflags(1) As Byte

Private m_textenc As Byte
Private m_sdata As String

Private Property Get clsF_frameid() As String
clsF_frameid = m_sframeid

End Property

Private Property Let clsF_frameid(ByVal a_sframeid As String)
m_sframeid = a_sframeid
    
End Property

Private Sub clsF_getflags(ByRef a_flags() As Byte)
a_flags = m_aflags
    
End Sub

Private Sub clsF_setflags(ByRef a_flags() As Byte)
Dim idx As Integer
For idx = 0 To UBound(m_aflags) 'zero based array
    m_aflags(idx) = a_flags(idx)
Next idx
    
End Sub

Private Property Get clsF_framesize() As Long
    clsF_framesize = Len(m_textenc) + Len(m_sdata)
End Property

Private Property Get clsF_textencoding() As Byte
    clsF_textencoding = m_textenc
End Property

Private Property Let clsF_textencoding(ByVal a_enc As Byte)
    m_textenc = a_enc
End Property

Private Property Get clsF_data() As String
    clsF_data = m_sdata
End Property

Private Property Let clsF_data(ByVal a_sdata As String)
    m_sdata = IIf(Right$(a_sdata, 1) <> vbNullChar, a_sdata & vbNullChar, a_sdata)
End Property

Private Sub clsF_getbytes(ByRef a_bytes() As Byte)
Dim idx As Integer

    ReDim a_bytes(clsF_framesize - 1)

    a_bytes(0) = m_textenc
    
    For idx = 1 To Len(m_sdata)
        a_bytes(idx) = Asc(Mid$(m_sdata, idx, 1))
    Next idx
    
End Sub

Private Sub clsF_parsebytes(ByRef a_bytes() As Byte)
Dim idx As Integer
m_textenc = a_bytes(0)

m_sdata = ""

For idx = 1 To UBound(a_bytes)
    m_sdata = m_sdata & Chr$(a_bytes(idx))
Next idx
    
End Sub

Private Sub class_initialize()
m_textenc = 0
m_sdata = vbNullChar

m_aflags(0) = 0
m_aflags(1) = 0

End Sub
