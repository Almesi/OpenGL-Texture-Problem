VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_VertexArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private p_ID As Long
Private p_VertexBuffer As std_Buffer
Private p_VertexLayout As std_BufferLayout

Private Sub Class_Initialize()
    Call glGenVertexArrays(1, p_ID)
End Sub

Private Sub Class_Terminate()
    Call glDeleteVertexArrays(1, p_ID)
End Sub

Public Function Create() As std_VertexArray
    Set Create = New std_VertexArray
    With Create
        .VertexBuffer = n_VertexBuffer
        .VertexLayout = n_VertexLayout
        .AddBuffer
    End With
End Function

Public Property Let ID(n_ID As Long)
    p_ID = n_ID
End Property
Public Property Get ID() As Long
    ID = p_ID
End Property
Public Property Get VertexBuffer() As std_Buffer
    Set VertexBuffer = p_VertexBuffer
End Property
Public Property Let VertexBuffer(n_VertexBuffer As std_Buffer)
    Set p_VertexBuffer = n_VertexBuffer
End Property

Public Property Get VertexLayout() As std_BufferLayout
    Set VertexLayout = p_VertexLayout
End Property
Public Property Let VertexLayout(n_VertexLayout As std_BufferLayout)
    Set p_VertexLayout = n_VertexLayout
End Property

Public Sub Bind()
    Call glBindVertexArray(p_ID)
End Sub

Public Sub UnBind()
    Call glBindVertexArray(0)
End Sub

Public Sub AddBuffer(n_VertexBuffer As std_Buffer, n_VertexLayout As std_BufferLayout)
    Dim Offset As Long
    Dim i As Long

    VertexBuffer = n_VertexBuffer
    VertexLayout = n_VertexLayout
    Call Bind()
    Call VertexBuffer.Bind()
    For i = 0 To Ubound(VertexLayout.Elements)
        With VertexLayout.Element(i)
            Call glEnableVertexAttribArray(i)
            Call glVertexAttribPointer(i, .Count, .Typee, .Normalized, VertexLayout.Stride(), Offset)
            Offset = Offset + (.Count * VertexLayout.GetSizeOfType(.Typee))
        End With
    Next i
End Sub