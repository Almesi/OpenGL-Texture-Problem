# OpenGL-Texture-Problem
This Temporary Rep is only here until i can solve the Texture Problem in the README File

## What am i trying to achieve?
I am a big fan of games. One of my hobbies is making games in Excel.
I am at a point in Game Development, where Excel just isnt cutting it anymore in terms of graphics rendering.
For that reason I am trying to create a Library for VBA-OpenGL.
A Collection of that already exists under https://arkham46.developpez.com/articles/office/vbaopengl/?page=page_1
This link has everything one might need to accomplish anything they want with OpenGL in VBA.
BUT, since it is OpenGL there you need lots of Code to see anything on the Screen. I am trying to create a Library with Classes that abstract the OpenGL code away.
So far it is going great.
Once i finish std_Texture i should have everything i need to finish the Library.

### I do have 2 Problems:
1. Problem is, i cannot use Indexbuffers. For some reason it doesnt work.
  My current solution is just declaring more Vertices. Basically doing what the IndexBuffer tries to solve-->Creating more Vertices.
2. Second Problem without a solution and the main reason i made this Repository is i cant get a Texture on the Screen

Some things i have tried:
Creating a Global Array to load the data. The idea is, what if the locally declared Array gets destroyed and the data points to nothing?
Creating the texture but not loading any data (using null pointer) and then assigning data later.

1. I tried glGetTexImage to check if the data was loaded correctly and it does.
2. I tried changing the vertex x/y/z position and the tx/ty position
3. I tried changing the Fragmenshader to make any color of value 0 to a specific color to check if the texture is just dark or completely empty(It was completely empty, every single pixel)
  Since everything is black that should mean, that the data was not assigned properly, but since glGetTexImage returns that data i dont know how to continue.
