# clsSharedMemory
Pass text between your applications.


First application:

Create instance:

  Private ShareMem As New clsSharedMemory
  
Open shared memory:

 ShareMem.Open("sharename")
 
Put text to memory:

 ShareMem.Put("text to pass")

 Other application:


Create instance:

  Private ShareMem As New clsSharedMemory
  
Open shared memory:

  ShareMem.Open("sharename")
  
Read text from memory:

  If ShareMem.DataExists Then Return ShareMem.Peek() 
      
I would like to pass on my experience in VB.Net to others and thus support this language for future generations. There are difficult beginnings in any programming language, when you have to learn to use new language libraries so that you can take even the smallest step. I want to help you with that now. You're welcome, signed Zdeněk Jantač.
