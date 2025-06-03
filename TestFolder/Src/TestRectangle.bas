Attribute VB_Name = "TestRectangle"


Option Explicit

Public VB             As std_Buffer
Public IB             As std_Buffer
Public VBLayout       As std_BufferLayout
Public VA             As std_VertexArray
Public Shader         As std_Shader
Public Texture        As std_Texture
Public Renderer       As std_Renderer
Public Window         As std_Window
Public MeshPositions  As std_Mesh
Public MeshIndices    As std_Mesh
Public FinalMesh      As std_Mesh

Public Const th256 As Single = 0.00390625

Public Function RunMain(Path As String) As Long

    #If VBA7 Then
        If LoadLibrary(Path & "\Freeglut64.dll") = False Then
            Debug.Print "Couldnt load freeglut"
            Exit Function
        End If
    #Else
        If LoadLibrary(Path & "\Freeglut.dll") = False Then
            Debug.Print "Couldnt load freeglut"
            Exit Function
        End If
    #End If

    Call glutInit(0&, "")
    Set Window = New std_Window
    Call Window.Create(1600, 900, GLUT_RGBA, "OpenGL Test", "4_6", GLUT_CORE_PROFILE, GLUT_DEBUG)

    Call GLStartDebug()

    Call glEnable(GL_BLEND)
    Call glEnable(GL_DEPTH_TEST)
    Call glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA)

    Call glFrontFace(GL_CW)

    Dim Color(11) As Single
    Color(00) = 1.0!: Color(01) = 0.5!: Color(02) = 0.5!
    Color(03) = 0.5!: Color(04) = 1.0!: Color(05) = 0.5!
    Color(06) = 1.0!: Color(07) = 1.0!: Color(08) = 0.5!
    Color(09) = 0.5!: Color(10) = 0.5!: Color(11) = 1.0!

    Dim Textures(7) As Single
    Textures(00) = 0.0!: Textures(01) = 0.0!
    Textures(02) = 1.0!: Textures(03) = 0.0!
    Textures(04) = 1.0!: Textures(05) = 1.0!
    Textures(06) = 0.0!: Textures(07) = 1.0!

    Set MeshPositions = std_Mesh.CreateStandardMesh(std_MeshType.Rectangle)
    Call MeshPositions.AddAttribute(3, Color)
    Call MeshPositions.AddAttribute(2, Textures)
    Set MeshIndices = std_Mesh.CreateStandardMeshIndex(std_MeshType.Rectangle)
    Set FinalMesh = std_Mesh.CreateMeshFromIndex(MeshPositions, MeshIndices)

    Set Shader = std_Shader.CreateFromFile(Path & "\Vertex.Shader", Path & "\Fragment.Shader")
    Set Texture = std_Texture.Create(Path & "\TestTexture3.png", GL_RGBA)
    Set VA = New std_VertexArray
    VA.Bind
    Set VB = std_Buffer.Create(GL_ARRAY_BUFFER, FinalMesh)
    Set IB = std_Buffer.Create(GL_ELEMENT_ARRAY_BUFFER, MeshIndices)

    Set VBLayout = New std_BufferLayout
    Call VBLayout.AddFloat(std_BufferLayoutType.XYZ)
    Call VBLayout.AddFloat(std_BufferLayoutType.RedGreenBlue)
    Call VBLayout.AddFloat(std_BufferLayoutType.TextureXTextureY)
    
    
    Call Texture.Bind()
    Call VA.AddBuffer(VB, VBLayout)
    Set Renderer = New std_Renderer

    Set IB = Nothing ' Remove line to test IndexBuffer

    Call glutDisplayFunc(AddressOf DrawLoop)
    Call glutIdleFunc(AddressOf DrawLoop)

    Call glutMainLoop

End Function

Public Sub DrawLoop()
    Static Count As Single
    Static Direction As Single
    Dim VertexColorLocation As Long
    Dim TextureLocation As Long
    Dim UniformName(8) As Byte
    UniformName(0) = Asc("o")
    UniformName(1) = Asc("u")
    UniformName(2) = Asc("r")
    UniformName(3) = Asc("C")
    UniformName(4) = Asc("o")
    UniformName(5) = Asc("l")
    UniformName(6) = Asc("o")
    UniformName(7) = Asc("r")
    ' Since VBA strings are 2 bytes per char i have to do this

    Dim UniformName2(8) As Byte
    UniformName2(0) = Asc("T")
    UniformName2(1) = Asc("e")
    UniformName2(2) = Asc("x")
    UniformName2(3) = Asc("t")
    UniformName2(4) = Asc("u")
    UniformName2(5) = Asc("r")
    UniformName2(6) = Asc("e")
    UniformName2(7) = Asc("0")
    ' Since VBA strings are 2 bytes per char i have to do this

    If Count >= 1 Then Direction = -th256
    If Count =< 0 Then Direction = +th256

    Call Shader.Bind
    VertexColorLocation = glGetUniformLocation(Shader.ID, VarPtr(UniformName(0)))
    Call glUniform4f(VertexColorLocation, 1.0!, 1.0!, 1.0!, 1.0!)
    TextureLocation = glGetUniformLocation(Shader.ID, VarPtr(UniformName2(0)))
    Call glUniform1i(TextureLocation, 0)
    Call Texture.Bind()



    Count = Count + Direction
    Call Renderer.Clear(Count * 2, Count * 3, Count, 1.0!)
    Call Renderer.Draw(VA, IB, Shader)
    Call glutSwapBuffers
End Sub