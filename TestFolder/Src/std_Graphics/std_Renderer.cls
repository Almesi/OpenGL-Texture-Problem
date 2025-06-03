VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Renderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Public Sub Clear(Optional r As Single = 0, Optional g As Single = 0, Optional b As Single = 0, Optional a As Single = 0)
    Call glClear(GL_COLOR_BUFFER_BIT OR GL_DEPTH_BUFFER_BIT)
    Call glClearColor(r, g, b, a)
End Sub

Public Sub Draw(VA As std_VertexArray, IndexBuffer As std_Buffer, Shader As std_Shader)
    Dim Count As Long
    
    Call VA.Bind()
    Call Shader.Bind()
    If IndexBuffer Is Nothing Then
        Count = VA.VertexBuffer.Mesh.Count / VA.VertexBuffer.Mesh.VertexSize
        Call glDrawArrays(GL_TRIANGLES, 0, Count)
    Else
        Call IndexBuffer.Bind()
        Call glDrawElements(GL_TRIANGLES, IndexBuffer.Mesh.Count, GL_UNSIGNED_INT, 0)
    End If
End Sub