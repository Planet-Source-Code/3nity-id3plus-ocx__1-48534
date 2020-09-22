VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWF"
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
Private m_sdescription As String
Private m_sdata As String

Public Property Get description() As String
Attribute description.VB_Description = "Return/Set the Description Field"
description = m_sdescription

End Property

Public Property Let description(ByVal a_sdesc As String)
m_sdescription = IIf(Right(a_sdesc, 1) <> vbNullChar, a_sdesc & vbNullChar, a_sdesc)

End Property

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
For idx = 0 To UBound(m_aflags)
    m_aflags(idx) = a_flags(idx)
Next idx
    
End Sub

Private Property Get clsF_framesize() As Long
clsF_framesize = Len(m_textenc) + Len(m_sdescription) + Len(m_sdata)

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
m_sdata = IIf(Right(a_sdata, 1) <> vbNullChar, a_sdata & vbNullChar, a_sdata)

End Property

Private Sub clsF_getbytes(ByRef a_bytes() As Byte)
Dim idx  As Integer
Dim idatastartpos As Integer

ReDim a_bytes(clsF_framesize - 1)

a_bytes(0) = m_textenc

For idx = 1 To Len(m_sdescription)
    a_bytes(idx) = Asc(Mid$(m_sdescription, idx, 1))
Next idx

idatastartpos = Len(m_textenc) + Len(m_sdescription)

For idx = 1 To Len(m_sdata)
    a_bytes(idatastartpos) = Asc(Mid$(m_sdata, idx, 1))
    idatastartpos = idatastartpos + 1
Next idx

End Sub

Private Sub clsF_parsebytes(ByRef a_bytes() As Byte)
Dim idx As Integer
Dim inullpos As Integer

m_textenc = a_bytes(0)

m_sdescription = ""

For idx = 1 To UBound(a_bytes)
    If Chr$(a_bytes(idx)) <> vbNullChar Then
        m_sdescription = m_sdescription & Chr$(a_bytes(idx))
    Else
        inullpos = idx
        m_sdescription = m_sdescription & vbNullChar
        Exit For
    End If
Next idx

m_sdata = ""

For idx = inullpos + 1 To UBound(a_bytes)
    m_sdata = m_sdata & Chr$(a_bytes(idx))
Next idx
    
End Sub

Private Sub class_initialize()
m_textenc = 0
m_sdescription = vbNullChar
m_sdata = vbNullChar

m_aflags(0) = 0
m_aflags(1) = 0
    
End Sub


