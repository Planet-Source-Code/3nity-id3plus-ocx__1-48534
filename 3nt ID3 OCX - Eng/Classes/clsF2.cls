VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsF2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colframes As Collection

Public Sub add(ByRef a_objframe As clsF)
Attribute add.VB_Description = "Add a Frame to the collection"

On Error GoTo errhandler
    
Call m_colframes.add(a_objframe)
    
normalexit:
On Error GoTo 0
Exit Sub

errhandler:
'Err.Raise Err.Number, Err.Source, Err.description

End Sub

Public Function item(ByVal a_index As Variant)
Attribute item.VB_Description = "Return a specified Frame"
    Set item = m_colframes.item(a_index)
End Function

Public Function count() As Long
Attribute count.VB_Description = "Return the number of Frames in the collection"
    count = m_colframes.count
End Function

Public Sub remove(ByVal a_index As Variant)
Attribute remove.VB_Description = "Remove a specified Frame"
    Call m_colframes.remove(a_index)
End Sub

Public Sub clear()
Attribute clear.VB_Description = "Remove all frames from the collection"

    Do While m_colframes.count > 0
        Call m_colframes.remove(1)
    Loop
    
End Sub

Private Sub class_initialize()
    Set m_colframes = New Collection
End Sub

Private Sub class_terminate()
    Set m_colframes = Nothing
End Sub

Public Function newenum() As IUnknown
Attribute newenum.VB_UserMemId = -4
Attribute newenum.VB_MemberFlags = "40"
    Set newenum = m_colframes.[_NewEnum]
End Function


