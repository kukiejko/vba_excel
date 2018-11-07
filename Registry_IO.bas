Attribute VB_Name = "Registry_IO"
Option Explicit


Sub TestRegistry()
Dim myRegKey As String
Dim myValue As String
Dim myAnswer As Integer


'demo at HKEY_CURRENT_USER\Software\EpiCollector\DB\
  'get registry key to work with
  myRegKey = InputBox("Which registry key do you want to read?", _
             "Get Registry Key")
  If myRegKey = "" Then Exit Sub
  'check if key exists
  If RegKeyExists(myRegKey) = True Then
    'key exists, read it
    myValue = RegKeyRead(myRegKey)
    'display result and ask if it should be changed
    myAnswer = MsgBox("The registry value for the key """ & _
               myRegKey & """" & vbCr & "is """ & myValue & _
               """" & vbCr & vbCr & _
               "Do you want to change it?", vbYesNo)
  Else
    'key doesn't exist, ask if it should be created
    myAnswer = MsgBox("The registry key """ & myRegKey & _
               """ could not be found." & vbCr & vbCr & _
               "Do you want to create it?", vbYesNo)
  End If
  If myAnswer = vbYes Then
    'ask for new registry key value
    myValue = InputBox("Please enter new value:", _
              myRegKey, myValue)
    If myValue <> "" Then
      'save/create registry key with new value
      RegKeySave myRegKey, myValue
      MsgBox "Registry key saved."
    End If
  End If
  
  'ask if key should be deleted from registry
  myAnswer = MsgBox("Do you want to delete the registry key """ & _
             myRegKey & """?", vbYesNo)
  If myAnswer = vbYes Then
    'delete registry key
    If RegKeyDelete(myRegKey) = True Then
      'deletion was successful
      MsgBox "Registry key """ & myRegKey & """ deleted."
    Else
      'deletion wasn't successful
      MsgBox "Registry key """ & myRegKey & _
             """ could not be deleted."
    End If
  End If
End Sub


Function RegKeyRead(i_RegKey As String) As String
Dim myWS As Object

  On Error Resume Next
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'read key from registry
  RegKeyRead = myWS.RegRead(i_RegKey)
End Function

Function RegKeyExists(i_RegKey As String) As Boolean
Dim myWS As Object

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'try to read the registry key
  myWS.RegRead i_RegKey
  'key was found
  RegKeyExists = True
  Exit Function
  
ErrorHandler:
  'key was not found
  RegKeyExists = False
End Function

Sub RegKeySave(i_RegKey As String, _
               i_Value As String, _
      Optional i_Type As String = "REG_SZ")
Dim myWS As Object

  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'write registry key
  myWS.RegWrite i_RegKey, i_Value, i_Type

End Sub

'deletes i_RegKey from the registry
'returns True if the deletion was successful,
'and False if not (the key couldn't be found)
Function RegKeyDelete(i_RegKey As String) As Boolean
Dim myWS As Object

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'delete registry key
  myWS.RegDelete i_RegKey
  'deletion was successful
  RegKeyDelete = True
  Exit Function

ErrorHandler:
  'deletion wasn't successful
  RegKeyDelete = False
End Function


