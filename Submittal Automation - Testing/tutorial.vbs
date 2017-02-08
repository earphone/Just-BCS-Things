  ' Input Box with only Prompt
  InputBox("Enter a number")    
  
  ' Input Box with a Title
  a=InputBox("Enter a Number","Enter Value")
  msgbox a
  
  ' Input Box with a Prompt,Title and Default value
  a=InputBox("Enter a Number","Enter Value",123)
  msgbox a
  
  ' Input Box with a Prompt,Title,Default and XPos
  a=InputBox("Enter your name","Enter Value",123,700)
  msgbox a
  
  ' Input Box with a Prompt,Title and Default and YPos
  a=InputBox("Enter your name","Enter Value",123,,500)
  msgbox a