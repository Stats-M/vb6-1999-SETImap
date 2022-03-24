Attribute VB_Name = "modErrorHandling"
Option Explicit

' Define your custom errors here.  Be sure to use numbers
' greater than 512, to avoid conflicts with OLE error numbers.
Public Const MyObjectError1000 = 1000
Public Const MyObjectError2 = 1010
Public Const MyObjectErrorN = 1234
Public Const MyUnhandledError = 9999

'Мои собственные константы
'Константы для заголовков MsgBox
Public Const StandartErrHeader = 1100
Public Const DataEntryErrHeader = 1101


Private Function GetErrorHeaderFromResource(ErrHeaderNum As Long) As String
On Error GoTo GetErrorHeaderFromResourceError

    GetErrorHeaderFromResource = LoadResString(ErrHeaderNum)

GetErrorHeaderFromResourceError:
      If Err.Number <> 0 Then
            GetErrorHeaderFromResource = "An unknown error has occurred!"
      End If
End Function


Private Function GetErrorTextFromResource(ErrorNum As Long) _
          As String
      Dim strMsg As String
      

      ' this function will retrieve an error description from a resource
      ' file (.RES).  The ErrorNum is the index of the string
      ' in the resource file.  Called by RaiseError


      On Error GoTo GetErrorTextFromResourceError
      

      ' get the string from a resource file
      GetErrorTextFromResource = LoadResString(ErrorNum)
      

      Exit Function
      

GetErrorTextFromResourceError:
      

      If Err.Number <> 0 Then
            GetErrorTextFromResource = "An unknown error has occurred!"
      End If
      

End Function


Public Sub RaiseError(ErrorNumber As Long, Source As String)
      Dim strErrorText As String


      'there are a number of methods for retrieving the error
      'message.  The following method uses a resource file to
      'retrieve strings indexed by the error number you are
      'raising.
      strErrorText = GetErrorTextFromResource(ErrorNumber)


      'raise an error back to the client
      Err.Raise vbObjectError + ErrorNumber, Source, strErrorText


End Sub

'**********************************************************
'*  Показывает сообщение об ошибках (текст - из ресурсов) *
'*  Начало нумерации - с 1200                             *
'**********************************************************

Public Sub RaiseErrMsg(ErrorNumber As Long, ErrHeader As Long)
Dim strErrorText As String
Dim strErrorHeader As String
      
    strErrorText = GetErrorTextFromResource(ErrorNumber)
    strErrorHeader = GetErrorHeaderFromResource(ErrHeader)
    Result = MsgBox(strErrorText, vbOKOnly, strErrorHeader)
End Sub
