#version 460 core
out vec4 FragColor;
  
in vec4 vertexcolor;
uniform vec4 ourColor;
in vec2 TexCoord;

uniform sampler2D Texture0;

void main()
{
    FragColor = texture(Texture0, TexCoord) + ourColor * vertexcolor;
}