VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Texture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Any, ByVal Length As Long)

Private p_ID As Long
Private p_FilePath As String
Private p_Data() As Byte
Private p_Width As Long
Private p_Height As Long
Private p_BPP As Long

Public Property Let ID(n_ID                  As Long)   : p_ID       = n_ID          : End Property
Public Property Let FilePath(n_FilePath      As String) : p_FilePath = n_FilePath    : End Property
Public Property Let Data(n_Data()            As Byte)   : p_Data     = n_Data        : End Property
Public Property Let Width(n_Width            As Long)   : p_Width    = n_Width       : End Property
Public Property Let Height(n_Height          As Long)   : p_Height   = n_Height      : End Property
Public Property Let BPP(n_BPP                As Long)   : p_BPP      = n_BPP         : End Property

Public Property Get ID()                     As Long    : ID         = p_ID          : End Property
Public Property Get FilePath()               As String  : FilePath   = p_FilePath    : End Property
Public Property Get Data()                   As Byte()  : Data       = p_Data        : End Property
Public Property Get Width()                  As Long    : Width      = p_Width       : End Property
Public Property Get Height()                 As Long    : Height     = p_Height      : End Property
Public Property Get BPP()                    As Long    : BPP        = p_BPP         : End Property

Public Property Get Pointer() As LongPtr
    Pointer = VarPtr(p_Data(0))
End Property


Public Function Create(Path As String, Format As Long) As std_Texture
    Dim TempID As Long
    Set Create = New std_Texture
    With Create
        Dim Image       As stdImage : Set Image     = stdImage.CreateFromFile(Path)
        Dim ColorData() As Long     :     ColorData = SwapColors(Image.Colors(), 1, 2, 3, 0) 'ARGB --> 'RGBA
        Dim ArrSize     As Long     :     ArrSize   = (Ubound(ColorData) + 1) * LenB(ColorData(1))

        Dim NewData() As Byte
        ReDim NewData(ArrSize - 1)
        Call CopyMemory(NewData(0), VarPtr(ColorData(0)), ArrSize)
        .Data     = NewData
        .Width    = 64 ' temporary
        .Height   = 64 ' temporary
        .BPP      = 4  ' temporary
        .FilePath = Path

        Call glGenTextures(1, TempID)
        .ID = TempID
        Call .Bind()
        Call glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR)
        Call glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR)
        Call glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_REPEAT)
        Call glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_REPEAT)
        Call glTexImage2D(GL_TEXTURE_2D, 0, Format, .Width, .Height, 0, Format, GL_UNSIGNED_BYTE, NewData(0))
    End With
End Function

Public Function SwapColors(Arr() As Long, ParamArray Order() As Variant) As Long()
    Dim i As Long, j As Long
    Dim ReturnArr() As Long
    Dim NewSize As Long
    Dim Temp() As Long
    Dim Element As Variant
    NewSize = Ubound(Arr, 1) * Ubound(Arr, 2) - 1
    ReDim ReturnArr(NewSize)
    ReDim Temp(UBound(Order))

    For Each Element In Arr
        ReturnArr(i) = Element
        i = i + 1
    Next Element

    For i = 0 To NewSize Step +4
        For j = 0 To Ubound(Order)
            Temp(j) = ReturnArr(i + Order(j))
        Next j
        For j = 0 To Ubound(Order)
            ReturnArr(i + j) = Temp(j)
        Next j
    Next i
    SwapColors = ReturnArr
End Function

Public Sub Bind()
    Call glActiveTexture(GL_TEXTURE0 + ID)
    Call glBindTexture(GL_TEXTURE_2D, ID)
End Sub

Public Sub Unbind()
    Call glBindTexture(GL_TEXTURE_2D, 0)
End Sub

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Call glDeleteTextures(1, ID)
End Sub