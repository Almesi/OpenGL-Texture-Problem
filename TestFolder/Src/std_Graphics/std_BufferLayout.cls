VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_BufferLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Size As Long
Private p_Stride As Long
Private p_Elements() As std_BufferElement

Public Enum std_BufferLayoutType
    XY               = 0
    XYZ              = 1
    RedGreenBlue     = 2
    TextureXTextureY = 3
End Enum

Public Property Let Stride(n_Stride As Long)
    p_Stride = n_Stride
End Property
Public Property Get Stride() As Long
    Stride = p_Stride
End Property

Public Property Let Elements(n_Elements() As std_BufferElement)
    Elements = p_Elements
End Property
Public Property Get Elements() As std_BufferElement()
    Elements = p_Elements
End Property

Public Property Let Element(Index As Long, n_Element As std_BufferElement)
    Set p_Elements(Index) = n_Element
End Property
Public Property Get Element(Index As Long) As std_BufferElement
    Set Element = p_Elements(Index)
End Property

Public Sub AddDouble(LayoutType As std_BufferLayoutType)      : Call Push(GL_DOUBLE       , LayoutType, GL_FALSE) : End Sub
Public Sub AddFloat(LayoutType As std_BufferLayoutType)       : Call Push(GL_FLOAT        , LayoutType, GL_FALSE) : End Sub
Public Sub AddUnsignedInt(LayoutType As std_BufferLayoutType) : Call Push(GL_UNSIGNED_INT , LayoutType, GL_FALSE) : End Sub
Public Sub AddUnsignedByte(LayoutType As std_BufferLayoutType): Call Push(GL_UNSIGNED_BYTE, LayoutType, GL_TRUE)  : End Sub

Public Sub Push(Typee As Long, LayoutType As std_BufferLayoutType, Normalized As Byte)
    Dim Count As Long

    Select Case LayoutType
        Case XY                : Count = 2
        Case XYZ               : Count = 3
        Case RedGreenBlue      : Count = 3
        Case TextureXTextureY  : Count = 2
    End Select

    Size = Size + 1
    ReDim Preserve p_Elements(Size)
    Element(Size) = std_BufferElement.Create(Typee, Count, Normalized)
    Stride = Stride + (Count * GetSizeOfType(Typee))
End Sub

Public Function GetSizeOfType(Typee As Long) As Long
    Select Case Typee
        Case GL_DOUBLE        : GetSizeOfType = 8
        Case GL_FLOAT         : GetSizeOfType = 4
        Case GL_UNSIGNED_INT  : GetSizeOfType = 4
        Case GL_UNSIGNED_BYTE : GetSizeOfType = 1
        Case Else             : GetSizeOfType = 0
    End Select
End Function

Private Sub Class_Initialize()
    Size = -1
End Sub