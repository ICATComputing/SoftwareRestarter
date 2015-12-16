'------------------------------------------------------------------------------------------------
' Filename    : modCommonFunctions.vb
' Purpose     : This is the common module that provides generic functions 
' Created By  : Felix Kang - I-CAT Computing (28 JUL 2005)
' Note        : 
' Assumptions : - Code is based on Visual Basic .NET (Visual Studio 2003)
'               - System.Drawing is added as reference
'               - System.Windows.Forms is added as reference
'------------------------------------------------------------------------------------------------
' History
' - 28 JUL 2005 : Creation date of the module
'------------------------------------------------------------------------------------------------

#Region " System Imports "

'Imports all the components we need
Imports System
Imports System.Data
Imports System.Xml
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Management

#End Region

Module modCommonFunctions

#Region " Constants "

  'Private Constants
  Private Const DEF_APP_SETTING_DS_NAME = "APP_SETTINGS"
  Private Const DEF_APP_SETTING_TABLENAME = "APP_SETTINGS"

  '.NET Data type keyword
  Public Const DOT_NET_STRING_KEYWORD = "System.String"
  Public Const DOT_NET_BOOLEAN_KEYWORD = "System.Boolean"
  Public Const DOT_NET_BYTE_KEYWORD = "System.Byte"
  Public Const DOT_NET_CHAR_KEYWORD = "System.Char"
  Public Const DOT_NET_DATETIME_KEYWORD = "System.DateTime"
  Public Const DOT_NET_DECIMAL_KEYWORD = "System.Decimal"
  Public Const DOT_NET_DOUBLE_KEYWORD = "System.Double"
  Public Const DOT_NET_SINGLE_KEYWORD = "System.Single"

  'Miscellaneous
  Private Const SRCCOPY As Integer = &HCC0020  
  Private WithEvents a As Printing.PrintDocument

#End Region

#Region " DLL definition "

Declare Auto Function BitBlt Lib "gdi32.dll" (ByVal _
  hdcDest As IntPtr, ByVal nXDest As Integer, ByVal _
  nYDest As Integer, ByVal nWidth As Integer, ByVal _
  nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc _
  As Integer, ByVal nYSrc As Integer, ByVal dwRop As _
  System.Int32) As Boolean

#End Region

#Region " Local Variables "

#End Region

#Region " Procedures / Functions "

Public Sub GenericErrorHandler(ByVal strFormName As String, ByVal strModuleName As String, ByVal lngErrorNo As Long, _
  ByVal strErrDesc As String, Optional ByVal strSQL As String = "", Optional ByVal strExtra As String = "", Optional ByVal strLogFilename As String = "")
'----------------------------------------------------------------------------------------------
' Purpose     : Initiliase the necessary components to start this application
' Assumption  : 
' Input       :
'   - strFormName, a string consist of the form/module where the error originated or detected
'   - strFunctionName, a string consist of function name where the error originated or detected
'   - lngErrorNo, a long variable to pass on any error number that gets generated
'   - strErrDesc, a string containing error description that gets generated
'   - strSQL (optional), a string to pass any SQL statement if this module gets called while executing SQL statement
'   - strExtra (optional), a string to pass any extra information that might be helpful in debugging process
'   - strLogFilename (optional), a string containing the full path and filename of error logfile. If this parameter is not
'     empty then error will be written to a file not screen. If nothing is passed then any error will be printed to screen
' Output      :
' Note        :
'   - At the moment error handler simply reports what's wrong, the calling function/procedures is responsible in cleaning up
'     or take the appropriate action.
'----------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim strTemp As String


  'Check whether a filename was passed (must contain full and valid filename (including path))
  If strLogFilename <> "" Then
    'Write it to file then
  Else
    'Print it to screen
    MsgBox("An application error has occured, please write down when, how it happened and the information below, " & _
    "to the appropriate support staff" & vbCrLf & vbCrLf & _
    "Form Name : " & strFormName & vbCrLf & _
    "Module Name : " & strModuleName & vbCrLf & _
    "Error Number : " & lngErrorNo & vbCrLf & _
    "Error Description : " & strErrDesc & vbCrLf & _
    "SQL Statement : " & strSQL & vbCrLf & _
    "Extra : " & strExtra, MsgBoxStyle.Critical, "Application error")
  End If

  Exit Sub

ErrorHandling:
  'To handle weird occassion where the error logging module itself has problem
  'Print it to screen
  MsgBox("An application error has occured, please write down when, how it happened and the information below, " & _
  "to the appropriate support staff" & vbCrLf & vbCrLf & _
  "Form Name : " & strFormName & vbCrLf & _
  "Module Name : " & strModuleName & vbCrLf & _
  "Error Number : " & lngErrorNo & vbCrLf & _
  "Error Description : " & strErrDesc & vbCrLf & _
  "SQL Statement : " & strSQL & vbCrLf & _
  "Extra : " & strExtra, MsgBoxStyle.Critical, "Application error")
End Sub

Public Function DoesFileExist(ByVal strFilepath As String) As Boolean
'------------------------------------------------------------------------------------------------
' Purpose     : Find out whether a file exist in the path specified
' Assumption  : 
' Input       :
'   - strFilepath, a string containing the complete path of the filename
' Output      :
'   - Boolean value, TRUE if the file exist and FALSE if it doesn't
' Created By  : Felix Kang
' Note        : 
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim fileInfo As IO.FileInfo


  'Create new instance of the FileInfo object
  fileInfo = New IO.FileInfo(strFilepath)
  'Check whether the file exist and return the info
  Return fileInfo.Exists

  Exit Function

ErrorHandling:
  'Report it to the user
  GenericErrorHandler("modCommonFunctions.vb", "DoesFileExist", Err.Number, Err.Description)
End Function

Public Function DoesFolderExist(ByVal strFolderPath As String) As Boolean
'------------------------------------------------------------------------------------------------
' Purpose     : Find out whether a file exist in the path specified
' Assumption  : 
' Input       :
'   - strFilepath, a string containing the complete path of the filename
' Output      :
'   - Boolean value, TRUE if the file exist and FALSE if it doesn't
' Created By  : Felix Kang
' Note        : 
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim dirInfo As IO.DirectoryInfo


  'Create new instance of the FileInfo object
  dirInfo = New IO.DirectoryInfo(strFolderPath)
  'Check whether the file exist and return the info
  Return dirInfo.Exists

  Exit Function

ErrorHandling:
  'Report it to the user
  GenericErrorHandler("modCommonFunctions.vb", "DoesFolderExist", Err.Number, Err.Description)
End Function

Public Function SaveSettingsInXML(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String, _
  ByVal strFilePath As String) As Boolean
'------------------------------------------------------------------------------------------------
' Purpose     : Saves application settings into an XML file
' Assumption  : 
' Input       :
'   - strSection, a string describing which section the setting belongs to
'   - strKey, a string describing which key the setting belongs to
'   - strValue, a string containing the actual value of the setting.
'   - strFilePath, a string with the complete + file name of the configuration file
' Output      :
'   - Generated/Updated XML file containing the settings
' Created By  : Felix Kang
' Note        : This function basically replaces INI file to locally store the application settings in XML file. 
'               The format of the settings however still mimic INI file with the Section, Key, Value format.
'               For those who comment it is an overkill to use DataSet to accomplish this, I agree with you, BUT
'               the effort needed to allow add,edit,update in XML file creates a huge amount of code, just for that.
'               Using the dataset's WriteXml makes the code shorter and allow possibilities for more complex structure.
'               If you can create a simple function to replace what I do below without adding too much code, be my guest :)
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim blnFileExist As Boolean
  Dim dsTempDBSettings As DataSet
  Dim dsTempSettingsTable As DataTable
  Dim dsTempDataRow() As DataRow
  Dim dsTempNewRow As DataRow


  'Check whether the old file exists
  blnFileExist = DoesFileExist(strFilePath)
  'Create a new dataset instance
  dsTempDBSettings = New DataSet
  If blnFileExist Then
    'If file exists, open the xml file containing the settings and load it into a dataset    
    dsTempDBSettings.ReadXml(strFilePath)
  Else
    'Create the necessary structure inside
    dsTempDBSettings.DataSetName = DEF_APP_SETTING_DS_NAME
    'Create a temp table
    dsTempSettingsTable = New DataTable
    dsTempSettingsTable.TableName = DEF_APP_SETTING_TABLENAME
    'Add the 'SECTION' Column
    dsTempSettingsTable.Columns.Add("SECTION", System.Type.GetType("System.String"))
    'Add the 'KEY' Column
    dsTempSettingsTable.Columns.Add("KEY", System.Type.GetType("System.String"))
    'Add the 'VALUE' Column
    dsTempSettingsTable.Columns.Add("VALUE", System.Type.GetType("System.String"))
    'Add this newly created table to the dataset
    dsTempDBSettings.Tables.Add(dsTempSettingsTable)
  End If

  'Now that we have a settings DB in the memory, look for the key that we want
  'NOTE: Hardcoded index of 0 has been given because we know that we only deal with 1 table
  dsTempDataRow = dsTempDBSettings.Tables(0).Select("SECTION='" & strSection & "' AND KEY='" & strKey & "'")
  'Check whether we have something
  If dsTempDataRow.Length < 1 Then
    'Create a new row if we don't find anything
    dsTempNewRow = dsTempDBSettings.Tables(0).NewRow
    'Assign the values
    dsTempNewRow("SECTION") = strSection
    dsTempNewRow("KEY") = strKey
    dsTempNewRow("VALUE") = strValue
    'Attach this to the collection of rows
    dsTempDBSettings.Tables(0).Rows.Add(dsTempNewRow)
  Else
    'If we do have it, just modify the existing value
    dsTempDataRow(0).Item("VALUE") = strValue
  End If

  'Once we're done, save it
  dsTempDBSettings.AcceptChanges()
  'And then save it as XML
  dsTempDBSettings.WriteXml(strFilePath)

  'Return TRUE to say everything is fine
  SaveSettingsInXML = True

  Exit Function

ErrorHandling:
  'Report it to the user
  GenericErrorHandler("modCommonFunctions.vb", "SaveSettingsInXML", Err.Number, Err.Description)
  SaveSettingsInXML = False
End Function

Public Function ReadSettingsInXML(ByVal strSection As String, ByVal strKey As String, ByVal strFilePath As String, _
  Optional ByRef strDefaultValue As String = "", Optional ByVal blnReturnNULLIfEmpty As Boolean = False) As String
'------------------------------------------------------------------------------------------------
' Purpose     : Saves application settings into an XML file
' Assumption  : 
' Input       :
'   - strSection, a string describing which section the setting belongs to
'   - strKey, a string describing which key the setting belongs to
'   - strValue, a string containing the actual value of the setting.
'   - strFilePath, a string with the complete + file name of the configuration file
'   - strDefaultValue, a different default value can be returned other than ""
' Output      :
'   - The value of the setting
' Created By  : Felix Kang
' Note        : This function basically replaces INI file to locally store the application settings in XML file. 
'               The format of the settings however still mimic INI file with the Section, Key, Value format.
'               For those who comment it is an overkill to use DataSet to accomplish this, I agree with you, BUT
'               the effort needed to allow add,edit,update in XML file creates a huge amount of code, just for that.
'               Using the dataset's WriteXml makes the code shorter and allow possibilities for more complex structure.
'               If you can create a simple function to replace what I do below without adding too much code, be my guest :)
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim blnFileExist As Boolean
  Dim dsTempDBSettings As DataSet
  Dim dsTempSettingsTable As DataTable
  Dim dsTempDataRow() As DataRow
  Dim dsTempNewRow As DataRow
  Dim strTemp As String


  'Check whether the old file exists
  blnFileExist = DoesFileExist(strFilePath)
  'Create a new dataset instance
  dsTempDBSettings = New DataSet
  If blnFileExist Then
    'If file exists, open the xml file containing the settings and load it into a dataset    
    dsTempDBSettings.ReadXml(strFilePath)
  Else
    'If the file doesn't even exist, why bother continuing
    If strDefaultValue = "" Then
      ReadSettingsInXML = "<NULL>"
    Else
      ReadSettingsInXML = strDefaultValue
    End If
    Exit Function
  End If

  'If we get this far, that means everything is ok, now that we have the DB in the memory, 
  'let's look for the key that we want
  'NOTE: Hardcoded index of 0 has been given because we know that we only deal with 1 table
  dsTempDataRow = dsTempDBSettings.Tables(0).Select("SECTION='" & strSection & "' AND KEY='" & strKey & "'")
  'Check whether we have something
  If dsTempDataRow.Length > 0 Then
    'If we have, just return the value
    strTemp = dsTempDataRow(0).Item("VALUE")
  Else
    ReadSettingsInXML = strDefaultValue
    Exit Function
  End If

  'Check to make sure we don't return empty value- if user doesn't want to
  If Trim(strTemp = "") Then
    'If it's empty, check whether user specify something for the default value
    If strDefaultValue = "" Then
      'If user doesn't specify anything, check whether they want NULL or not
      If blnReturnNULLIfEmpty = True Then
        ReadSettingsInXML = "<NULL>"
      Else
        ReadSettingsInXML = ""
      End If
      'Exit from here
      Exit Function
    Else
      'Return the default value
      ReadSettingsInXML = strDefaultValue
    End If
  Else
    'If what we get is not empty - return the value
    ReadSettingsInXML = strTemp
  End If

  Exit Function

ErrorHandling:
  'Report it to the user
  GenericErrorHandler("modCommonFunctions.vb", "ReadSettingsInXML", Err.Number, Err.Description)
  ReadSettingsInXML = ""
End Function

Public Function DeleteSettingsInXML(ByVal strFilePath As String, ByVal strSection As String, _
  Optional ByVal strKey As String = "") As Boolean
'------------------------------------------------------------------------------------------------
' Purpose     : Delete application settings from an XML file
' Assumption  : 
' Input       :
'   - strFilePath, a string with the complete + file name of the configuration file
'   - strSection, a string describing which section the setting belongs to
'   - strKey (optional), a string describing which key the setting belongs to
' Output      :
'   - a TRUE if it was successful, FALSE if it was fail
' Created By  : Felix Kang
' Note        : This function basically replaces INI file to locally store the application settings in XML file. 
'               The format of the settings however still mimic INI file with the Section, Key, Value format.
'               For those who comment it is an overkill to use DataSet to accomplish this, I agree with you, BUT
'               the effort needed to allow add,edit,update in XML file creates a huge amount of code, just for that.
'               Using the dataset's WriteXml makes the code shorter and allow possibilities for more complex structure.
'               If you can create a simple function to replace what I do below without adding too much code, be my guest :)
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim blnFileExist As Boolean
  Dim dsTempDBSettings As DataSet
  Dim dsTempSettingsTable As DataTable
  Dim dsTempDataRow() As DataRow
  Dim dsTempNewRow As DataRow
  Dim intRowLoop As Integer


  'Check whether the old file exists
  blnFileExist = DoesFileExist(strFilePath)
  If blnFileExist = False Then
    'No need to continue if even the files doesn't exist
    DeleteSettingsInXML = False
    Exit Function
  End If

  'If we're here that means we're good to go - create a new dataset instance
  dsTempDBSettings = New DataSet
  'Open the xml file containing the settings and load it into a dataset    
  dsTempDBSettings.ReadXml(strFilePath)

  'If user specified a key
  If Trim(strKey) <> "" Then
    'Search the section and key that we want to delete
    'NOTE: Hardcoded index of 0 has been given because we know that we only deal with 1 table
    dsTempDataRow = dsTempDBSettings.Tables(0).Select("SECTION='" & strSection & "' AND KEY='" & strKey & "'")
  Else
    'Search the section that we want to delete
    'NOTE: Hardcoded index of 0 has been given because we know that we only deal with 1 table
    dsTempDataRow = dsTempDBSettings.Tables(0).Select("SECTION='" & strSection & "'")
  End If

  'Delete them
  For intRowLoop = 0 To dsTempDataRow.Length - 1
    dsTempDataRow(intRowLoop).Delete()
  Next

  'Once we're done, save it
  dsTempDBSettings.AcceptChanges()
  'And then save it as XML
  dsTempDBSettings.WriteXml(strFilePath)

  'Return TRUE to say everything was smooth
  DeleteSettingsInXML = True

  Exit Function

ErrorHandling:
  'Report it to the user
  GenericErrorHandler("modCommonFunctions.vb", "DeleteSettingsInXML", Err.Number, Err.Description)
  DeleteSettingsInXML = False
End Function

Public Function RSAEncrypt(ByVal strInputString As String, ByVal strPublicKey As String, Optional ByVal blnSucccess As Boolean = False) As String
'------------------------------------------------------------------------------------------------
' Purpose     : Encrypt a piece of string using RSA strength Encryption
' Assumption  : 
' Input       :
'   - strInputString, the string that will be encrypted
'   - strPublicKey, the RSA public key needed to do the encyption, generated elsewhere
' Output      :
'   - The encrypted string
' Created By  : Felix Kang
' Note        : The limitation of encrypting this is, the input string cannot be longer than 58 characters. 
'               If you have more than 58 characters to process, split it into a block of 58 characters.
'               The unused variables xmlPrivateKey and xmlPublicKey are there so developer can quickly generate a pair key if needed
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim cryRSACryptoProvider As RSACryptoServiceProvider
  Dim xmlPrivateKey As String
  Dim xmlPublicKey As String
  Dim bytPlainTextBArray As Byte()
  Dim bytCypherTextBArray As Byte()
  Dim intLoopCounter As Integer
  Dim strTempArray As String()
  Dim strOutputString As String


  'Create a new instance of the crypto provider
  cryRSACryptoProvider = New RSACryptoServiceProvider
  'Create Private Key
  xmlPrivateKey = cryRSACryptoProvider.ToXmlString(True)
  'Create Public Key
  xmlPublicKey = cryRSACryptoProvider.ToXmlString(False)

  'Check that we have no more than 58 char
  If strInputString.Length > 58 Then
    'Generate error and quit
    RSAEncrypt = "ERR:Length more than 58"
    blnSucccess = False
    Exit Function
  End If

  'Get the public key so message can be encrypted
  cryRSACryptoProvider.FromXmlString(strPublicKey)
  'Transform message string into an array of bytes
  ReDim bytPlainTextBArray(strInputString.Length - 1)
  For intLoopCounter = 0 To strInputString.Length - 1
    bytPlainTextBArray(intLoopCounter) = CByte(Asc(Mid(strInputString, intLoopCounter + 1, 1)))
  Next
  'bytPlainTextBArray = (New UnicodeEncoding).GetBytes(strInputString)
  'Encrypt 
  bytCypherTextBArray = cryRSACryptoProvider.Encrypt(bytPlainTextBArray, False)
  'Convert it to a string and move them into a separate container
  ReDim strTempArray(bytCypherTextBArray.Length - 1)
  For intLoopCounter = 0 To bytCypherTextBArray.Length - 1
    'Create the temp container
    strTempArray(intLoopCounter) = bytCypherTextBArray(intLoopCounter).ToString
  Next
  'And then join them into one long string
  strOutputString = Join(strTempArray, " ")

  'Pass the value back
  RSAEncrypt = strOutputString

  Exit Function

ErrorHandling:
  'Report it to the user
  GenericErrorHandler("modCommonFunctions.vb", "RSAEncrypt", Err.Number, Err.Description)
End Function

Public Function RSADecrypt(ByVal strEncryptedString As String, ByVal strRSAObjectBothKeys As String, Optional ByVal blnSucccess As Boolean = False) As String
'------------------------------------------------------------------------------------------------
' Purpose     : Saves application settings into an XML file
' Assumption  : 
' Input       :
' Output      :
'   - The value of the setting
' Created By  : Felix Kang
' Note        : The limitation of encrypting this is, the input string cannot be longer than 58 characters. 
'               If you have more than 58 characters to process, split it into a block of 58 characters.
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim cryRSACryptoProvider As RSACryptoServiceProvider
  Dim xmlPrivateKey As String
  Dim xmlPublicKey As String
  Dim bytPlainTextBArray As Byte()
  Dim bytCypherTextBArray As Byte()
  Dim intLoopCounter As Integer
  Dim strOutputString As String
  Dim strTempArray As String()


  'Create a new instance of the crypto provider
  cryRSACryptoProvider = New RSACryptoServiceProvider
  'Get the RSA Object key
  cryRSACryptoProvider.FromXmlString(strRSAObjectBothKeys)

  'Split the string to an array elements
  strTempArray = Split(strEncryptedString, " ")
  'Transform it back to an array of bytes so it can be decrypted
  ReDim bytCypherTextBArray(strTempArray.Length - 1)
  For intLoopCounter = 0 To strTempArray.Length - 1
    bytCypherTextBArray(intLoopCounter) = CByte(strTempArray(intLoopCounter))
  Next intLoopCounter
  'Decrypt it
  bytPlainTextBArray = cryRSACryptoProvider.Decrypt(bytCypherTextBArray, False)
  'Rebuild the string
  For intLoopCounter = 0 To bytPlainTextBArray.Length - 1
    strOutputString = strOutputString & Chr(bytPlainTextBArray(intLoopCounter))
  Next intLoopCounter

  'Pass the value
  RSADecrypt = strOutputString

  Exit Function

ErrorHandling:
  'Report it to the user
  GenericErrorHandler("modCommonFunctions.vb", "RSADecrypt", Err.Number, Err.Description)
End Function

Public Function GetFilesListIntoArray(ByVal strPath As String, Optional ByVal strFileExtension As String = "*.*", _
  Optional ByVal intFileCount As Integer = 0) As String()
'------------------------------------------------------------------------------------------------
' Purpose     : This gets the list of files in a specific directory
' Assumption  : 
' Input       :
'   - strPath, string containing path where the files are
'   - strFileExtension, string containing file extension if only specific file are needed, wildcard format (*.*, *.txt, etc)
'   - intFileCount, integer with the file count
' Output      :
'   - a zero based array of string, containing all of the file matching the criteria. 
'     If there's no file it will return "Nothing"
' Created By  : Felix Kang
' Note        : 
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim diTargetFolder As DirectoryInfo
  Dim fiFileList As FileInfo()
  Dim fiIndividualFile As FileInfo
  Dim intTempCount As Integer
  Dim strTempArray As String()


  'Create a new instance to point to the folder
  diTargetFolder = New DirectoryInfo(strPath)
  'Get the file list based on extension criteria
  fiFileList = diTargetFolder.GetFiles(strFileExtension)

  'Get the file count
  intTempCount = fiFileList.Length()

  'If there's nothing, no need to go further
  If intTempCount <= 0 Then
    Exit Function
  End If

  'If there's something, then we need to redefine the destination array
  ReDim strTempArray(intTempCount - 1)

  'Spits out the filename one by one and put it into Array of strings
  intTempCount = 0
  For Each fiIndividualFile In fiFileList
    'Move the filename one by one
    strTempArray(intTempCount) = fiIndividualFile.Name
    'Increment the array pointer
    intTempCount = intTempCount + 1
  Next

  'Once it is finish, return the value
  GetFilesListIntoArray = strTempArray

  Exit Function

ErrorHandling:
  'Report it to the user
  GenericErrorHandler("modCommonFunctions.vb", "GetFilesListIntoArray", Err.Number, Err.Description)
End Function

Public Function ConvertAmountHoursToDecimal(ByVal strAmountHours As String, ByVal strDelimiter As String) As Decimal
'------------------------------------------------------------------------------------------------
' Purpose     : This gets the list of files in a specific directory
' Assumption  : 
' Input       :
'   - strAmountHours, a string with format "HH:MM" (Hours:Minutes)
'   - strDelimiter, a string defining the delimiter
' Output      :
'   - a decimal converting the the based 60 of hours into based 100 (decimal)
' Created By  : Felix Kang
' Note        : 
'   - Please note that this is NOT a time converter. It convert an amount of hours and minutes which formatted as "HH:MM". 
'     as a result the hours could very well be more than 24
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim strHoursComponent As String()
  Dim intHours As Integer
  Dim decMinutes As Decimal
  Dim decFinalFigure As Decimal
  Dim intTemp As Integer


  'If we have blank or zero in the string and no delimiter, return zero
  If (Trim(strAmountHours) = "") Then
    ConvertAmountHoursToDecimal = 0
    Exit Function
  End If

  'Check if it has a delimiter
  If InStr(strAmountHours, strDelimiter) < 1 Then
    'if it doesn't, try to convert it into a number - if we can do that in the first place)
    If IsNumeric(strAmountHours) = True Then
      ConvertAmountHoursToDecimal = CDbl(strAmountHours)
      Exit Function
    Else
      ConvertAmountHoursToDecimal = -9999
      Exit Function
    End If
  End If

  'Split HH:MM format into its components which are hours and minutes (zero based arrays)
  strHoursComponent = Split(strAmountHours, strDelimiter)

  'Get the hours, it's safe to assume that the 1st portion of the array will be the hour
  If Trim(strHoursComponent(0)) <> "" Then
    intHours = CInt(strHoursComponent(0))
  Else
    intHours = 0
  End If
  'Get the minutes, it's safe to assume that the 2nd portion of the array will be the minutes
  If Trim(strHoursComponent(1)) <> "" Then
    intTemp = CInt(strHoursComponent(1))
    'Convert the minutes into a decimal form
    decMinutes = intTemp / 60
  Else
    decMinutes = 0
  End If

  'Return the value
  ConvertAmountHoursToDecimal = CDec(intHours) + decMinutes

  Exit Function

ErrorHandling:
  ConvertAmountHoursToDecimal = -9999
  'Report it to the user
  GenericErrorHandler("modCommonFunctions.vb", "ConvertAmountHoursToDecimal", Err.Number, Err.Description)
End Function

Public Function ExtractFileNameOnly(ByVal strCompleteFileName As String) As String
'------------------------------------------------------------------------------------------------
' Purpose     : This extract the path portion of a complete filename
' Assumption  : 
'   - legitimate filename input includes extension
' Input       :
'   - strCompleteFileName, a string containing the complete filename
' Output      :
'   - a string containing just the filename, minus path, and minus extension
' Created By  : Felix Kang
' Note        : 
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim fiSourceFile As FileInfo


  'Create a new instance that points to the source file
  fiSourceFile = New FileInfo(strCompleteFileName)
  'Return the filename only
  ExtractFileNameOnly = Replace(fiSourceFile.Name, fiSourceFile.Extension, "")  'Replace the extension with blank

  Exit Function

ErrorHandling:
  'Report it to the user
  GenericErrorHandler("modCommonFunctions.vb", "ExtractFileNameOnly", Err.Number, Err.Description)
  ExtractFileNameOnly = ""
End Function

Public Function ExtractPathOnly(ByVal strCompleteFileName As String) As String
'------------------------------------------------------------------------------------------------
' Purpose     : This extract the path portion of a complete filename
' Assumption  : 
'   - The filename input includes path
' Input       :
'   - strCompleteFileName, a string containing the complete filename, including the path
' Output      :
'   - a string containing just the path
' Created By  : Felix Kang
' Note        : 
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim fiSourceFile As FileInfo


  'Create a new instance that points to the source file
  fiSourceFile = New FileInfo(strCompleteFileName)
  'Return the filename only
  ExtractPathOnly = fiSourceFile.DirectoryName

  Exit Function

ErrorHandling:
  'Report it to the user
  GenericErrorHandler("modCommonFunctions.vb", "ExtractPathOnly", Err.Number, Err.Description)
  ExtractPathOnly = ""
End Function

Public Function ExtractFileNameWithExtension(ByVal strCompleteFileName As String) As String
'------------------------------------------------------------------------------------------------
' Purpose     : This function strips the path portion of a complete filename, leaving filename + extension
' Assumption  : 
'   - legitimate filename input includes extension
' Input       :
'   - strCompleteFileName, a string containing the complete filename
' Output      :
'   - a string containing just the filename, minus path, and minus extension
' Created By  : Felix Kang
' Note        : 
'------------------------------------------------------------------------------------------------
  On Error GoTo ErrorHandling

  Dim fiSourceFile As FileInfo


  'Create a new instance that points to the source file
  fiSourceFile = New FileInfo(strCompleteFileName)
  'Return the filename only
  ExtractFileNameWithExtension = fiSourceFile.Name

  Exit Function

ErrorHandling:
  'Report it to the user
  GenericErrorHandler("modCommonFunctions.vb", "ExtractFileNameWithExtension", Err.Number, Err.Description)
  ExtractFileNameWithExtension = ""
End Function

Public Function CharacterPadding(ByVal strRepeatChar As Char, ByVal intRepeatNumber As Integer) As String
'---------------------------------------------------------------------
'Purpose     : This procedure is to generate a string with repeating characters
'Assumptions :
'Input       :
'   - strRepeatChar, the char that will be repeated
'   - intRepeatNumber, the number of repeats
'Returns     :
'   - String containing the repeated character
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim strOutput As String

  'Initialise the string object (create new instance?) 
  strOutput = ""
  CharacterPadding = strOutput.PadLeft(intRepeatNumber, strRepeatChar)

  Exit Function
ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "CharacterPadding", Err.Number, Err.Source, Err.Description)
  CharacterPadding = False
End Function

Public Function ConvertColorToString(ByVal oColors As System.Drawing.Color) As String
'---------------------------------------------------------------------
'Purpose     : This procedure is to convert a color object into a string with format ARGB (A=,R=,G=,B=)
'Assumptions :
'Input       :
'   - oColors, the input colours
'Returns     :
'   - String containing the string representation of the input color
'Note        : 
'---------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim strTemp As String


  'Get the string representation of the input color
  strTemp = oColors.ToString
  'The string representation is in Color [A=,R=,G=,B=] so obviously it needs to be stripped into more manageable form
  'A=,R=,G=,B=
  ConvertColorToString = strTemp.Substring(strTemp.IndexOf("[") + 1, (((strTemp.LastIndexOf("]") - 1) - strTemp.IndexOf("["))))

  Exit Function

ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "ConvertColorToString", Err.Number, Err.Source, Err.Description)
  ConvertColorToString = "ERR"
End Function

Public Function ConvertStringToColor(ByVal strColors As String, Optional ByRef intAlphaComp As Integer = -1, _
  Optional ByRef intRedComp As Integer = -1, Optional ByRef intGreenComp As Integer = -1, _
  Optional ByRef intBlueComp As Integer = -1) As System.Drawing.Color
'---------------------------------------------------------------------
'Purpose     : This procedure is to convert a string with format ARGB (A=,R=,G=,B=) into color object
'Assumptions :
'Input       :
'   - oColors, the input colour in string format (A=,R=,G=,B=)
'   - intAlphaComp (optional), integer value of the Alpha component of the converted color
'   - intRedComp (optional), integer value of the Red component of the converted color
'   - intGreenComp (optional), integer value of the Green component of the converted color
'   - intBlueComp (optional), integer value of the Blue component of the converted color
'Returns     :
'   - the color object representation of the string
'Note        : 
'   - This function is meant to be used as tool to convert the string back from ConvertColorToString into a color object
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim strTempArray As String()
  Dim strTempAlpha As String
  Dim strTempRed As String
  Dim strTempGreen As String
  Dim strTempBlue As String
  Dim tempRGBColor As Integer


  'Check whether this is in A=,R=,G=,B= format or not
  If InStr(strColors, ",") > 0 Then
    'If we found something, we need to break up the string into its own component
    strTempArray = Split(strColors, ",")
    'Get the Alpha Component of the color
    strTempAlpha = Mid(strTempArray(0), InStr(strTempArray(0), "=") + 1)
    intAlphaComp = CInt(strTempAlpha)
    'Get the Red Component of the color
    strTempRed = Mid(strTempArray(1), InStr(strTempArray(1), "=") + 1)
    intRedComp = CInt(strTempRed)
    'Get the Green Component of the color
    strTempGreen = Mid(strTempArray(2), InStr(strTempArray(2), "=") + 1)
    intGreenComp = CInt(strTempGreen)
    'Get the Blue Component of the color
    strTempBlue = Mid(strTempArray(3), InStr(strTempArray(3), "=") + 1)
    intBlueComp = CInt(strTempBlue)
    'and once we got them all, convert them back into the color object
    ConvertStringToColor = ConvertStringToColor.FromArgb(intAlphaComp, intRedComp, intGreenComp, intBlueComp)
  Else
    If (Trim(strColors) = "") Or (strColors = "<NULL>") Then
      ConvertStringToColor = ConvertStringToColor.White
    Else
      'If not, we presume it's common/known color, convert the single color into RGB
      ConvertStringToColor = ConvertStringToColor.FromName(strColors)
    End If
  End If

  Exit Function
ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "ConvertStringToColor", Err.Number, Err.Source, Err.Description)
  ConvertStringToColor = Nothing
End Function

Public Function RecursiveFileList(ByVal strStartFolder As String, ByRef colSearchResult As Collection, _
  Optional ByVal strSearchCriteria As String = "*.*") As Boolean
'---------------------------------------------------------------------
'Purpose     : This procedure is to list all of the files under the start folder, recursively
'Assumptions :
'Input       :
'   - strStartFolder, a string specifying the folder to start the search
'   - colSearchResult, a collection of string to manage the search result, easier this way
'   - strSearchCriteria, a string specifying the file search criteria (i.e, *.*, *.tif, etc)
'Returns     :
'   - a Boolean value, TRUE if it went smooth, FALSE if it didn't
'Note        : 
'   - Core idea of the code is from http://www.bubble-media.com/cgi-bin/articles/archives/000026.html and it has been changed
'     to match our code style
'   - Recursive is memory intense and it should not be performed on a deep and complex directory structure
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim strDirectoryList() As String
  Dim strFileList() As String
  Dim intFileLoop As Integer
  Dim intDirLoop As Integer
  Dim blnSuccess As Boolean


  'Add trailing separators to the supplied paths if they don't exist.
  If Not strStartFolder.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
    strStartFolder = strStartFolder & System.IO.Path.DirectorySeparatorChar
  End If

  'We continue to drill down into directory structure.
  'Get a list of sub-directories from the current parent.
  strDirectoryList = System.IO.Directory.GetDirectories(strStartFolder)

  'For each directory listed, do a recursive
  For intDirLoop = 0 To strDirectoryList.GetUpperBound(0)
    blnSuccess = RecursiveFileList(strDirectoryList(intDirLoop), colSearchResult, strSearchCriteria)
  Next

  'Once it cannot go down any lower, we can start getting the files 
  strFileList = System.IO.Directory.GetFiles(strStartFolder, strSearchCriteria)
  'Loop and add the file into collection
  For intFileLoop = 0 To strFileList.GetUpperBound(0)
    'Add the file list into collection
    colSearchResult.Add(strFileList(intFileLoop))
  Next

  'If we get this far that means everything is ok
  RecursiveFileList = True

  Exit Function
ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "RecursiveFileList", Err.Number, Err.Source, Err.Description)
  RecursiveFileList = False
End Function

Public Function WriteTextToFile(ByVal strText As String, ByVal strFilename As String, ByVal blnAppend As Boolean) As Boolean
'---------------------------------------------------------------------
'Purpose     : This procedure is to list all of the files under the start folder, recursively
'Assumptions :
'Input       :
'   - strText, a string to write to the textfile
'   - strFilename, a string specifying the file to wrote to
'   - blnAppend, a boolean value stating whether the text should appending or overwriting
'Returns     :
'   - a Boolean value, TRUE if it went smooth, FALSE if it didn't
'Note        : 
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim stwTextStreamWrite As StreamWriter


  'Create new instance
  stwTextStreamWrite = New StreamWriter(strFilename, blnAppend)

  'Write the text into the destination file
  stwTextStreamWrite.Write(strText & vbCrLf)

  'Close the file
  stwTextStreamWrite.Close()

  'If we get this far, everything is fine
  WriteTextToFile = True

  Exit Function
ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "WriteTextToFile", Err.Number, Err.Source, Err.Description)
  WriteTextToFile = False
End Function

Public Sub KeystrokeAcceptCurrencyOnly(ByVal objKeystrokeValue As KeyEventArgs)
'---------------------------------------------------------------------
'Purpose     : This procedure restricts users keystroke for currency data entry only
'Assumptions :
'Input       :
'   - objKeystrokeValue, a KeyEventArgs objects passed down from TextBox KeyDown event
'Returns     :
'Note        : 
'   - What it does is quite simple, if it's a valid keyboard value for currency data entry then we don't do anything,
'     If it's not, we send a backspace key to the system, effectively deleting whatever they have entered
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim intKbdValue As Integer


  'Transfer the keyboard value to a variable for easy access
  intKbdValue = objKeystrokeValue.KeyValue
  'If it's part of the keys needed for currency data entry, let it go
  Select Case intKbdValue
    Case Keys.D0 To Keys.D9
      'Numeric value
    Case Keys.NumPad0 To Keys.NumPad9
      'Numeric value
    Case Keys.Decimal
      'decimal / dot
    Case Keys.OemPeriod
      'decimal / dot
    Case Keys.Back
      'Backspace
    Case Keys.Oemcomma
      'Comma
    Case Keys.OemMinus
      'Minus    
    Case Keys.Subtract
      'Minus
    Case 36
      '$ sign    
    Case Else
      'If it's not on the list, send a backspace
      SendKeys.Send("{BACKSPACE}")
  End Select

  Exit Sub
ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "KeystrokeAcceptCurrencyOnly", Err.Number, Err.Source, Err.Description)
End Sub

Public Sub KeystrokeAcceptNumbersOnly(ByVal objKeystrokeValue As KeyEventArgs)
'---------------------------------------------------------------------
'Purpose     : This procedure restricts users keystroke for numbers data entry only
'Assumptions :
'Input       :
'   - objKeystrokeValue, a KeyEventArgs objects passed down from TextBox KeyDown event
'Returns     :
'Note        : 
'   - What it does is quite simple, if it's a valid keyboard value for numbers data entry then we don't do anything,
'     If it's not, we send a backspace key to the system, effectively deleting whatever they have entered
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim intKbdValue As Integer


  'Transfer the keyboard value to a variable for easy access
  intKbdValue = objKeystrokeValue.KeyValue
  'If it's part of the keys needed for numbers data entry, let it go
  Select Case intKbdValue
    Case Keys.D0 To Keys.D9
      'Numeric value
    Case Keys.NumPad0 To Keys.NumPad9
      'Numeric value
    Case Keys.Back
      'Backspace
    Case Else
      'If it's not on the list, send a backspace
      SendKeys.Send("{BACKSPACE}")
  End Select

  Exit Sub
ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "KeystrokeAcceptNumbersOnly", Err.Number, Err.Source, Err.Description)
End Sub

Public Function StripNullCharacters(ByVal vstrStringWithNulls As String) As String
'---------------------------------------------------------------------
'Purpose     : Stripping NULL Characters
'Assumptions :
'Input       :
'Returns     :
'   - a clean string
'Note        : 
'   - Code base was taken from http://www.freevbcode.com/ShowCode.Asp?ID=4520
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim intPosition As Integer
  Dim strStringWithOutNulls As String


  intPosition = 1
  strStringWithOutNulls = vstrStringWithNulls

  Do While intPosition > 0
    intPosition = InStr(intPosition, vstrStringWithNulls, vbNullChar)

    If intPosition > 0 Then
      strStringWithOutNulls = Left$(strStringWithOutNulls, intPosition - 1) & _
      Right$(strStringWithOutNulls, Len(strStringWithOutNulls) - intPosition)
    End If

    If intPosition > strStringWithOutNulls.Length Then
      Exit Do
    End If
  Loop

  Return strStringWithOutNulls

  Exit Function
ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "StripNullCharacters", Err.Number, Err.Source, Err.Description)
  StripNullCharacters = ""
End Function

Public Function GetFormBitmapImage(ByVal frmTarget As Form) As Bitmap
'---------------------------------------------------------------------
'Purpose     : Create a bitmap image of the target form
'Assumptions :
'Input       :
' - frmTarget, a form that is going to be printed
'Returns     :
'   - a bitmap image
'Note        : 
'   - Code base was taken from http://www.vb-helper.com/howto_net_print_form_image.html
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim grFormTarget As Graphics
  Dim bmFormBitmap As Bitmap
  Dim grFormBitmap As Graphics
  Dim ipFormHDC As IntPtr
  Dim ipBitmapHDC As IntPtr

  'Get the form's graphic object(?)
  grFormTarget = frmTarget.CreateGraphics()
  'Create a bitmap container for the graphic(?)
  bmFormBitmap = New Bitmap(frmTarget.ClientSize.Width, frmTarget.ClientSize.Height, grFormTarget)
  'Create a graphic from the bitmap(?)
  grFormBitmap = grFormTarget.FromImage(bmFormBitmap)

  'Get a pointer to the bitmap object
  ipBitmapHDC = grFormBitmap.GetHdc()
  'Get a pointer from form's graphic object
  ipFormHDC = grFormTarget.GetHdc()

  'Transfer the data from form's image into bitmap (?) using BitBlt
  BitBlt(ipBitmapHDC, 0, 0, frmTarget.ClientSize.Width, frmTarget.ClientSize.Height, _
  ipFormHDC, 0, 0, SRCCOPY)

  'Release the pointers
  grFormBitmap.ReleaseHdc(ipBitmapHDC)
  grFormTarget.ReleaseHdc(ipFormHDC)

  'Return the result
  Return bmFormBitmap

  Exit Function
ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "GetFormBitmapImage", Err.Number, Err.Source, Err.Description)
  GetFormBitmapImage = Nothing
End Function

Public Function IsProcessRunning(ByVal strServerName As String, ByVal strProcessName As String) As Boolean
'---------------------------------------------------------------------
'Purpose     : Check whether a particular process is running or not
'Assumptions :
'Input       :
' - strServerName, a string that specify the server name where the process needs to be checked
' - strProcessName, a string that specify the process name to be checked
'Returns     :
'   - TRUE if the process is running, FALSE but it doesn't
'Note        : 
'   - Code base was taken from http://www.freevbcode.com/ShowCode.asp?ID=5166
'   - Requires WMI
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim objProcess As Object
  Dim strObjectString As String


  'Set the default to False
  IsProcessRunning = False
  'Compose the object string
  strObjectString = "winmgmts://" & strServerName

  'Check the process
  For Each objProcess In GetObject(strObjectString).InstancesOf("win32_process")
    If UCase(objProcess.name) = UCase(strProcessName) Then
      IsProcessRunning = True
      Exit Function
    End If
  Next

  Exit Function

ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "IsProcessRunning", Err.Number, Err.Source, Err.Description)
  IsProcessRunning = False
End Function

Public Function CheckExistingProcessInstance() As Boolean
'---------------------------------------------------------------------
'Purpose     : Check whether another copy of this program/process/instance is running or not
'Assumptions :
'Input       :
' - strProcessName, a string that specify the process name to be checked
'Returns     :
'   - TRUE if the process is running, FALSE but it doesn't
'Note        : 
'   - Code base was taken from http://www.freevbcode.com/ShowCode.asp?ID=5166
'   - Requires WMI
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandler


  'Get number of processes of you program
  If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then
    CheckExistingProcessInstance = True
  Else
    CheckExistingProcessInstance = False
  End If

  Exit Function

ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "CheckExistingProcess", Err.Number, Err.Source, Err.Description)
  CheckExistingProcessInstance = False
End Function

Public Function IsFileInUse(ByVal strFileName As String) As Boolean
'---------------------------------------------------------------------
'Purpose     : If the file is already opened by another process and the specified type of access is not allowed,
'              the Open operation fails and an error occurs.
'Assumptions :
'Input       : 
' - strFileName, a string specifying the file, complete with path and filename
'Returns     :
'   - TRUE if the the file in use and locked, FALSE if the file is not in use.
'Note        : 
'   - Code base was taken from http://www.thescripts.com/forum/thread627953.html
'-------------------------------------------------------------------------------    

  Dim objFile As System.IO.FileStream


  Try
    'Try to open the file
    objFile = New System.IO.FileStream(strFileName, FileMode.Open, FileAccess.Write, FileShare.None)
    'Close it
    objFile.Close()
    'Return the result
    isFileInUse = False
  Catch ex As Exception
    'If we get error message that means the file is open and in use
    isFileInUse = True
  End Try

End Function

Public Function RecursiveCloseFileList(ByVal strStartFolder As String, ByRef colSearchResult As Collection, _
  Optional ByVal strSearchCriteria As String = "*.*") As Boolean
'---------------------------------------------------------------------
'Purpose     : This procedure is to list all of the close files under the start folder, recursively. Close file means it doesn't
'              have an application locking into it so we are free to do whatever we want to it.
'Assumptions :
'Input       :
'   - strStartFolder, a string specifying the folder to start the search
'   - colSearchResult, a collection of string to manage the search result, easier this way
'   - strSearchCriteria, a string specifying the file search criteria (i.e, *.*, *.tif, etc)
'Returns     :
'   - a Boolean value, TRUE if it went smooth, FALSE if it didn't
'Note        : 
'   - Core idea of the code is from http://www.bubble-media.com/cgi-bin/articles/archives/000026.html and it has been changed
'     to match our code style
'   - Recursive is memory intense and it should not be performed on a deep and complex directory structure
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim strDirectoryList() As String
  Dim strFileList() As String
  Dim intFileLoop As Integer
  Dim intDirLoop As Integer
  Dim blnSuccess As Boolean


  'Add trailing separators to the supplied paths if they don't exist.
  If Not strStartFolder.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
    strStartFolder = strStartFolder & System.IO.Path.DirectorySeparatorChar
  End If

  'We continue to drill down into directory structure.
  'Get a list of sub-directories from the current parent.
  strDirectoryList = System.IO.Directory.GetDirectories(strStartFolder)

  'For each directory listed, do a recursive
  For intDirLoop = 0 To strDirectoryList.GetUpperBound(0)
    blnSuccess = RecursiveCloseFileList(strDirectoryList(intDirLoop), colSearchResult, strSearchCriteria)
  Next

  'Once it cannot go down any lower, we can start getting the files 
  strFileList = System.IO.Directory.GetFiles(strStartFolder, strSearchCriteria)
  'Loop and add the file into collection
  For intFileLoop = 0 To strFileList.GetUpperBound(0)
    'Add the file list into collection, only if the file is closed
    If IsFileInUse(strFileList(intFileLoop)) = False Then
      colSearchResult.Add(strFileList(intFileLoop))
    End If
  Next

  'If we get this far that means everything is ok
  RecursiveCloseFileList = True

  Exit Function
ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "RecursiveCloseFileList", Err.Number, Err.Source, Err.Description)
  RecursiveCloseFileList = False
End Function

Public Function GetHDSerialNumber() As String
'---------------------------------------------------------------------
'Purpose     : This function is to retrieve the HD manufacturer's serial number. One application is to generate license key.
'Assumptions :
'Input       :
'Returns     :
'   - a Boolean value, TRUE if it went smooth, FALSE if it didn't
'Note        : 
'   - Core idea of the code is from http://www.codeproject.com/csharp/hard_disk_serialNo.asp and it has been changed
'     to match our code style. We only returning the 1st 
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim objHWSearch As ManagementObjectSearcher
  Dim objWMI_HDObject As ManagementObject
  Dim strTempString As String
  Dim intTempLooop As Integer


  'Create a new instance of the object, we're searching in Win32_PhysicalMedia for the HD
  objHWSearch = New ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMedia")

  'Loop to get each record
  strTempString = ""
  For Each objWMI_HDObject In objHWSearch.Get
    'Check whether we get a NULL value
    If Trim(objWMI_HDObject("SerialNumber")) = vbNull.ToString Then
      'If yes, append NULL to the return string
      strTempString = strTempString & "NULL,"
    Else
      'Put it into a temp object placeholder
      strTempString = strTempString & Trim(objWMI_HDObject("SerialNumber")) & ","
    End If
  Next

  'Return the value - Minus the last comma
  GetHDSerialNumber = Mid(strTempString, 1, Len(strTempString) - 1)

  Exit Function
ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "GetHDSerialNumber", Err.Number, Err.Source, Err.Description)
  GetHDSerialNumber = "ERR"
End Function

Public Function GetMotherboardSerialNumber() As String
'---------------------------------------------------------------------
'Purpose     : This function is to retrieve the Motherboard's serial number. One application is to generate license key.
'Assumptions :
'Input       :
'Returns     :
'   - a Boolean value, TRUE if it went smooth, FALSE if it didn't
'Note        : 
'   - Core idea of the code is from http://www.codeproject.com/csharp/hard_disk_serialNo.asp and it has been changed
'     to match our code style. We only returning the 1st 
'-------------------------------------------------------------------------------
  On Error GoTo ErrorHandler

  Dim objHWSearch As ManagementObjectSearcher
  Dim objWMI_MBObject As ManagementObject
  Dim strTempString As String
  Dim intTempLooop As Integer


  'Create a new instance of the object, we're searching in Win32_BaseBoard for the motherboard
  objHWSearch = New ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard")

  'Loop to get each record
  strTempString = ""
  For Each objWMI_MBObject In objHWSearch.Get
    'Check whether we get a NULL value
    If Trim(objWMI_MBObject("SerialNumber")) = vbNull.ToString Then
      'If yes, append NULL to the return string
      strTempString = strTempString & "NULL"
    Else
      'Put it into a temp object placeholder
      strTempString = strTempString & Trim(objWMI_MBObject("SerialNumber"))
    End If
  Next

  'Return the value - we can assume most PCs have 1 motheboard
  GetMotherboardSerialNumber = strTempString

  Exit Function
ErrorHandler:
  GenericErrorHandler("modCommonFunctions.vb", "GetMotherboardSerialNumber", Err.Number, Err.Source, Err.Description)
  GetMotherboardSerialNumber = "ERR"
End Function

#End Region

End Module