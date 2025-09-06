' Declare functions
Declare Function returnNumUserInput () As double
Declare Function returnOprUserInput () As string
Declare Function returnCalcResult () As string
Declare Function isNumeric (Byval uInput as string) As boolean
Declare Function userQuit (Byval uInput as string) As boolean
Declare Function divideByZeroErr () As boolean
Declare Function continueCalcChk () as boolean

' Declare variables
Dim Shared As double userNumInput0, userNumInput1, result
Dim Shared As string userOprInput
Dim Shared as boolean contChk, alwaysCont

' Function to take number inputs
Function returnNumUserInput () As double
   Dim as string uInput
   Dim as double uInputResult
   Do
   print "Please enter a number."
   Input "", uInput
   If userQuit(uInput) then
      End
   end if
   If isNumeric(uInput) then
      uInputResult = val(uInput)
      return uInputResult
   else
      print "You have inputted text, when a number was expected!"
   end if
   Loop
end function

' Function to take operator inputs
Function returnOprUserInput () as string
   Dim as string uInput
   Do
   print "Please enter an operator."
   Input "", uInput
   If userQuit(uInput) then
      End
   end if
   If Len(uInput) = 1 AndAlso Instr(uInput, Any "+-*/") then
      return uInput
   else
      Print "You have inputted an invalid option."
   end if 
   Loop
end function

' Function to calculate the result of the number inputs
Function returnCalcResult () as string
   Select Case userOprInput
      Case "+"
         result = userNumInput0 + userNumInput1
      Case "-"
         result = userNumInput0 - userNumInput1
      Case "*"
         result = userNumInput0 * userNumInput1
      Case "/"
         result = userNumInput0 / userNumInput1
   End Select
   return userNumInput0 & " " & userOprInput & " " & userNumInput1 & " = " & result
end function

' Function to check whether or not the numerical input is valid
Function isNumeric (Byval uInput as string) as boolean
   Dim as integer dot = 1
   For index As integer = 0 To Len(uInput) - 1
      Select Case As Const uInput[index]
         case Asc("0") To Asc("9")
         Case Asc(".")
            If dot > 0 then
               dot -= 1
            else
               Return false
            end if
         Case Else
            return false
      end select
   next index
   return true
end function

' Function that allows user to quit early
Function userQuit (Byval uInput as string) As boolean
   uInput = LCase(uInput)
   If Left(uInput, 4) = "quit" then
      return true
   end if
   return false
end function

' Checks that the calculation does not contain a divide by zero error
Function divideByZeroErr () As boolean
   If userNumInput1 = 0 And userOprInput = "/" then
      Print "Divide by zero error!" + !"\n" + "Returning to start of calculations..."
      return true
   end if
   return false
end function

Function continueCalcChk () as boolean
   Dim as string uInput
   Do
   print "Do you wish to continue this calculation using previous results? Y/N/Always"
   Input "", uInput
   uInput = LCase(uInput)
   If userQuit(uInput) then
      End
   end if
   Select case uInput
      case "y", "yes"
         print "You have chosen to continue!"
         return true
      case "n", "no"
         return false
      case "a", "always"
         print "You have selected to always continue for this session." + !"\n" + "Reminder: you can quit at any time by typing 'quit'"
         alwaysCont = true
         return true
      case else
         print "You have input an invalid character/string!"
      End Select
      Loop
end function

'Main subroutine, calls all functions and does some basic logic checks
Sub Main()
   Do
   If contChk = false then
      print "Type 'quit' to quit."
      userNumInput0 = returnNumUserInput
   End if
   userOprInput = returnOprUserInput
   userNumInput1 = returnNumUserInput
   If divideByZeroErr() then
      Continue Do
   end if
   print returnCalcResult() 
   If alwaysCont orelse continueCalcChk() then
      userNumInput0 = result
      contChk = true
      Continue Do
   end if
   print "Press any key to quit..."
   sleep
   End
   Loop
end sub

Main()