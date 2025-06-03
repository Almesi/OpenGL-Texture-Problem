VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private p_ID As Long
Private p_BufferType As Long
Private p_Mesh As std_Mesh

Public Property Let ID(n_ID As Long)
    p_ID = n_ID
End Property

Public Property Get ID() As Long
    ID = p_ID
End Property

Public Property Get Mesh() As std_Mesh
    Set Mesh = p_Mesh
End Property
Public Property Let Mesh(n_Mesh As std_Mesh)
    Set p_Mesh = n_Mesh
End Property

Public Property Get BufferType() As Long
    BufferType = p_BufferType
End Property
Public Property Let BufferType(n_BufferType As Long)
    p_BufferType = n_BufferType
End Property

Public Function Create(n_BufferType As Long, n_Mesh As std_Mesh) As std_Buffer
    Dim TempID As Long
    Set Create = New std_Buffer
    With Create
        .BufferType = n_BufferType
        .Mesh       = n_Mesh
        Call glGenBuffers(1, TempID)
        .ID = TempID
        Call glBindBuffer(n_BufferType, .ID)
        Call glBufferData(n_BufferType, .Mesh.ByteSize, .Mesh.Pointer, GL_STATIC_DRAW)
    End With
End Function

Public Sub Bind()
    Call glBindBuffer(p_BufferType, p_ID)
End Sub

Public Sub Unbind()
    Call glBindBuffer(p_BufferType, 0)
End Sub

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Call glDeleteBuffers(1, p_ID)
End Sub