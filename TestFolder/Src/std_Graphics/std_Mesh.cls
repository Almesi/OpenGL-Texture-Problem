VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Mesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Any, ByVal Length As Long)

Private Const pi As Single = 3.14159

Private p_Data As Variant
Private p_ByteData() As Byte
Private p_Typee As Long
Private p_Count As Long
Private p_Size As Long
Private p_VertexSize As Long

Public Enum std_MeshType
    NoPoint    = 00
    Point      = 01
    Line       = 02
    Triangle   = 03
    Rectangle  = 04
    Pentagon   = 05
    Hexagon    = 06
    Variable   = 07
    Pyramid    = 08
    Box        = 09
    Variable3D = 10
End Enum

Public Property Let Data(n_Data              As Variant) : p_Data       = n_Data       : End Property
Public Property Let Typee(n_Typee As Long)               : p_Typee      = n_Typee      : End Property
Public Property Let Count(n_Count As Long)               : p_Count      = n_Count      : End Property
Public Property Let Size(n_Size As Long)                 : p_Size       = n_Size       : End Property
Public Property Let VertexSize(n_VertexSize  As Long)    : p_VertexSize = n_VertexSize : End Property 

Public Property Get Data()                   As Variant  : Data         = p_Data       : End Property
Public Property Get Typee()                  As Long     : Typee        = p_Type       : End Property
Public Property Get Count()                  As Long     : Count        = p_Count      : End Property
Public Property Get Size()                   As Long     : Size         = p_Size       : End Property
Public Property Get VertexSize()             As Long     : VertexSize   = p_VertexSize : End Property
Public Property Get DataIndex(Index As Long) As Single   : DataIndex    = p_Data(Index): End Property

Public Property Get ByteData() As Byte()
    ByteData = p_ByteData
End Property

Public Function ByteSize() As Long
    ByteSize = Ubound(ByteData()) + 1
End Function

Public Property Get Pointer() As LongPtr
    Pointer = VarPtr(p_ByteData(0))
End Function

Public Sub SetByteData()
    Dim Temp() As Byte
    Dim TempArr() As Single
    TempArr = p_Data
    Dim NewSize As Long
    NewSize = (Ubound(Data) + 1) * Size
    ReDim Temp(NewSize - 1)
    Call CopyMemory(Temp(0), VarPtr(TempArr(0)), NewSize)
    p_ByteData = Temp
End Sub

Public Function Create(n_Data() As Single, n_Typee As Long, n_Count As Long, n_Size As Long, n_VertexSize As Long) As std_Mesh
    Set Create = New std_Mesh
    With Create
        .Data       = n_Data
        .Typee      = n_Typee
        .Count      = n_Count
        .Size       = n_Size
        .VertexSize = n_VertexSize
        .SetByteData
    End With
End Function

Public Function CreateStandardMesh(MeshType As std_MeshType, Optional VariableSize As Long = 0) As std_Mesh
    Dim MeshData() As Single
    Dim i As Long
    Dim NewSize  As Long
    Dim PiFactor As Single
    Dim Radian   As Single
    NewSize  = VariableSize * 3
    PiFactor = 3.14 / 180
    Select Case MeshType
        Case std_MeshType.NoPoint
        Case std_MeshType.Point
            ReDim MeshData(2)
        Case std_MeshType.Line
            ReDim MeshData(5)
            MeshData(0) = 0.0!: MeshData(1) = 0.0!: MeshData(2) = 0.0!
            MeshData(3) = 1.0!: MeshData(4) = 0.0!: MeshData(5) = 0.0!
        Case std_MeshType.Triangle
            ReDim MeshData(8)
            MeshData(0) = 0.0!: MeshData(1) = 0.0!: MeshData(2) = 0.0!
            MeshData(3) = 1.0!: MeshData(4) = 0.0!: MeshData(5) = 0.0!
            MeshData(6) = 1.0!: MeshData(7) = 1.0!: MeshData(8) = 0.0!
        Case std_MeshType.Rectangle
            ReDim MeshData(11)
            MeshData(00) = -0.5!: MeshData(01) = -0.5!: MeshData(02) = 0.0!
            MeshData(03) = +0.5!: MeshData(04) = -0.5!: MeshData(05) = 0.0!
            MeshData(06) = +0.5!: MeshData(07) = +0.5!: MeshData(08) = 0.0!
            MeshData(09) = -0.5!: MeshData(10) = +0.5!: MeshData(11) = 0.0!
        Case std_MeshType.Pentagon
            ReDim MeshData(14)
            MeshData(00) = +0.25!: MeshData(01) = -1.0!: MeshData(02) = +0.00!
            MeshData(03) = +1.00!: MeshData(04) = +0.0!: MeshData(05) = +0.00!
            MeshData(06) = +0.00!: MeshData(07) = +1.0!: MeshData(08) = +0.00!
            MeshData(09) = -1.00!: MeshData(10) = +0.0!: MeshData(11) = +0.00!
            MeshData(12) = -0.25!: MeshData(13) = -1.0!: MeshData(14) = +0.00!
        Case std_MeshType.Hexagon
            ReDim MeshData(17)
            MeshData(00) = +0.00!: MeshData(01) = -1.00!: MeshData(02) = +0.00!
            MeshData(03) = +1.00!: MeshData(04) = -0.25!: MeshData(05) = +0.00!
            MeshData(06) = +1.00!: MeshData(07) = +0.25!: MeshData(08) = +0.00!
            MeshData(09) = +0.00!: MeshData(10) = +1.00!: MeshData(11) = +0.00!
            MeshData(12) = -1.00!: MeshData(13) = +0.25!: MeshData(14) = +0.00!
            MeshData(15) = -1.00!: MeshData(16) = -0.25!: MeshData(17) = +0.00!
        Case std_MeshType.Variable
            If VariableSize < 3 Then Exit Function
            ReDim MeshData(NewSize - 1)
            For i = 1 To VariableSize
                Radian = i * (360 / VariableSize) * PiFactor
                MeshData(3 * i - 3    ) = CSng(Sin(Radian))
                MeshData(3 * i - 3 + 1) = CSng(Cos(Radian))
                MeshData(3 * i - 3 + 2) = 0
            Next i
        Case std_MeshType.Pyramid
            ReDim MeshData(14)
            MeshData(00) = +1.00!: MeshData(01) = +0.00!: MeshData(02) = -1.00!
            MeshData(03) = -1.00!: MeshData(04) = +0.00!: MeshData(05) = -1.00!
            MeshData(06) = +1.00!: MeshData(07) = +0.00!: MeshData(08) = +1.00!
            MeshData(09) = +1.00!: MeshData(10) = +0.00!: MeshData(11) = +1.00!
            MeshData(12) = +0.00!: MeshData(13) = +1.00!: MeshData(14) = +0.00!
        Case std_MeshType.Box
            ReDim MeshData(23)
            MeshData(00) = -1.00!: MeshData(01) = -1.00!: MeshData(02) = -1.00!
            MeshData(03) = +1.00!: MeshData(04) = -1.00!: MeshData(05) = -1.00!
            MeshData(06) = +1.00!: MeshData(07) = +1.00!: MeshData(08) = -1.00!
            MeshData(09) = -1.00!: MeshData(10) = +1.00!: MeshData(11) = -1.00!
            MeshData(12) = -1.00!: MeshData(13) = -1.00!: MeshData(14) = +1.00!
            MeshData(15) = +1.00!: MeshData(16) = -1.00!: MeshData(17) = +1.00!
            MeshData(18) = +1.00!: MeshData(19) = +1.00!: MeshData(20) = +1.00!
            MeshData(21) = -1.00!: MeshData(22) = +1.00!: MeshData(23) = +1.00!
        Case std_MeshType.Variable3D
            If VariableSize < 6 Then Exit Function
            'MeshData = SphereMesh(VariableSize)
    End Select
    Set CreateStandardMesh = Create(MeshData, vbSingle, Ubound(MeshData) + 1, LenB(MeshData(0)), 3)
End Function

Public Function CreateStandardMeshIndex(MeshType As std_MeshType, Optional VariableSize As Long = 0) As std_Mesh
    Dim MeshData() As Single
    Dim i As Long
    Dim NewSize  As Long
    NewSize  = (VariableSize - 2) * 3
    Select Case MeshType
        Case std_MeshType.NoPoint
        Case std_MeshType.Point
            ReDim MeshData(0)
        Case std_MeshType.Line
            ReDim MeshData(1)
            MeshData(0) = 0: MeshData(1) = 1
        Case std_MeshType.Triangle
            ReDim MeshData(2)
            MeshData(0) = 0: MeshData(1) = 1: MeshData(2) = 2
        Case std_MeshType.Rectangle
            ReDim MeshData(5)
            MeshData(00) = 0: MeshData(01) = 3: MeshData(02) =  2
            MeshData(03) = 0: MeshData(04) = 2: MeshData(05) =  1
        Case std_MeshType.Pentagon
            ReDim MeshData(8)
            MeshData(00) = 0: MeshData(01) = 2: MeshData(02) =  1
            MeshData(03) = 0: MeshData(04) = 3: MeshData(05) =  2
            MeshData(06) = 0: MeshData(07) = 4: MeshData(08) =  3
        Case std_MeshType.Hexagon
            ReDim MeshData(11)
            MeshData(00) = 0: MeshData(01) = 2: MeshData(02) = 1
            MeshData(03) = 0: MeshData(04) = 3: MeshData(05) = 2
            MeshData(06) = 0: MeshData(07) = 4: MeshData(08) = 3
            MeshData(09) = 0: MeshData(10) = 5: MeshData(11) = 4
        Case std_MeshType.Variable
            If VariableSize = < 3 Then Exit Function
            ReDim MeshData(NewSize - 1)
            For i = 0 To VariableSize - 1
                MeshData(3 * i    ) = 0
                MeshData(3 * i + 1) = i + 2
                MeshData(3 * i + 2) = i + 2 - 1
            Next i
        Case std_MeshType.Pyramid
            ReDim MeshData(17)
            MeshData(00) = 0: MeshData(01) = 3: MeshData(02) = 1
            MeshData(03) = 0: MeshData(04) = 2: MeshData(05) = 3
            MeshData(06) = 0: MeshData(07) = 4: MeshData(08) = 1
            MeshData(09) = 1: MeshData(10) = 4: MeshData(11) = 3
            MeshData(12) = 3: MeshData(13) = 4: MeshData(14) = 2
            MeshData(15) = 2: MeshData(16) = 4: MeshData(17) = 0
        Case std_MeshType.Box
            ReDim MeshData(35)
            MeshData(00) = 0: MeshData(01) =  2: MeshData(02) =  1
            MeshData(03) = 0: MeshData(04) =  3: MeshData(05) =  2
            MeshData(06) = 1: MeshData(07) =  2: MeshData(08) =  6
            MeshData(09) = 1: MeshData(10) =  6: MeshData(11) =  5
            MeshData(12) = 5: MeshData(13) =  6: MeshData(14) =  7
            MeshData(15) = 5: MeshData(16) =  7: MeshData(17) =  4
            MeshData(18) = 4: MeshData(19) =  7: MeshData(20) =  3
            MeshData(21) = 4: MeshData(22) =  3: MeshData(23) =  0
            MeshData(24) = 4: MeshData(25) =  0: MeshData(26) =  1
            MeshData(27) = 4: MeshData(28) =  1: MeshData(29) =  5
            MeshData(30) = 3: MeshData(31) =  7: MeshData(32) =  6
            MeshData(33) = 3: MeshData(34) =  6: MeshData(35) =  2
        Case std_MeshType.Variable3D
            If VariableSize < 6 Then Exit Function
            ' MeshData = SphereMesh(VariableSize)
    End Select
    Set CreateStandardMeshIndex = Create(MeshData, vbSingle, Ubound(MeshData) + 1, LenB(MeshData(0)), 3)
End Function

' AttributeData needs either as many elements as ItemsPerVertex, which will then be used for every input
' Or It has to have enough elements for each input as seperate data
Public Sub AddAttribute(ItemsPerVertex As Long, AttributeData() As Single)
    Dim i As Long, j As Long, k As Long
    Dim OldVertexSize As Long  : OldVertexSize = VertexSize
    Dim VertexCount   As Long  : VertexCount = (UBound(Data) + 1) / VertexSize

    Dim NewDataPointer As Long
    Dim DataPointer As Long
    Dim AttributePointer As Long

    VertexSize = VertexSize + ItemsPerVertex
    Dim TempData() As Single
    ReDim TempData(VertexCount * VertexSize - 1)
    For i = 0 To VertexCount - 1
        For j = 0 To (OldVertexSize - 1)
            TempData(NewDataPointer) = Data(DataPointer)
            NewDataPointer = NewDataPointer + 1
            DataPointer = DataPointer + 1
        Next j

        If Ubound(AttributeData) + 1 = ItemsPerVertex Then
            For k = 1 To ItemsPerVertex
                TempData(NewDataPointer) = AttributeData(AttributePointer)
                NewDataPointer = NewDataPointer + 1
            Next k
        Else
            For k = 1 To ItemsPerVertex
                TempData(NewDataPointer) = AttributeData(AttributePointer)
                AttributePointer = AttributePointer + 1
                NewDataPointer = NewDataPointer + 1
            Next k
        End If
    Next i
    Data = TempData
    Count = Ubound(TempData) + 1
End Sub

Public Function CreateMeshFromIndex(VertexMesh As std_Mesh, IndexMesh As std_Mesh) As std_Mesh
    Dim i As Long, j As Long
    Dim VertexData() As Single : VertexData = VertexMesh.Data()
    Dim IndexData()  As Single : IndexData  = IndexMesh.Data()
    Dim NewSize      As Long   : NewSize    = VertexMesh.VertexSize * (Ubound(IndexData) + 1) - 1
    Dim MeshData() As Single   : MeshData   = ReturnDataAsType(vbSingle, NewSize)

    For i = 0 To UBound(IndexData)
        For j = 0 To VertexMesh.VertexSize - 1
            MeshData(i * VertexMesh.VertexSize + j) = VertexData(IndexData(i) * VertexMesh.VertexSize + j)
        Next j
    Next i
    Set CreateMeshFromIndex = Create(MeshData, vbSingle, Ubound(MeshData) + 1, LenB(MeshData(0)), VertexMesh.VertexSize)
End Function

Public Function ReturnDataAsType(VariableType As vbVarType, Size As Long) As Variant
    Dim ReturnData As Variant
    Select Case VariableType
        Case vbBoolean  : ReturnData = CreateArrayBoolean(Size)
        Case vbInteger  : ReturnData = CreateArrayInteger(Size)
        Case vbLong     : ReturnData = CreateArrayLong(Size)
        Case vbLongLong : ReturnData = CreateArrayLongLong(Size)
        Case vbSingle   : ReturnData = CreateArraySingle(Size)
        Case vbDouble   : ReturnData = CreateArrayDouble(Size)
        Case vbString   : ReturnData = CreateArrayString(Size)
        Case vbByte     : ReturnData = CreateArrayByte(Size)
    End Select
    ReturnDataAsType = ReturnData
End Function

Private Function CreateArrayBoolean(Size  As Long) As Boolean()  : Dim Temp() As Boolean  : ReDim Temp(Size): CreateArrayBoolean  = Temp: End Function
Private Function CreateArrayInteger(Size  As Long) As Integer()  : Dim Temp() As Integer  : ReDim Temp(Size): CreateArrayInteger  = Temp: End Function
Private Function CreateArrayLong(Size     As Long) As Long()     : Dim Temp() As Long     : ReDim Temp(Size): CreateArrayLong     = Temp: End Function
Private Function CreateArrayLongLong(Size As Long) As LongLong() : Dim Temp() As LongLong : ReDim Temp(Size): CreateArrayLongLong = Temp: End Function
Private Function CreateArraySingle(Size   As Long) As Single()   : Dim Temp() As Single   : ReDim Temp(Size): CreateArraySingle   = Temp: End Function
Private Function CreateArrayDouble(Size   As Long) As Double()   : Dim Temp() As Double   : ReDim Temp(Size): CreateArrayDouble   = Temp: End Function
Private Function CreateArrayString(Size   As Long) As String()   : Dim Temp() As String   : ReDim Temp(Size): CreateArrayString   = Temp: End Function
Private Function CreateArrayByte(Size     As Long) As Byte()     : Dim Temp() As Byte     : ReDim Temp(Size): CreateArrayByte     = Temp: End Function