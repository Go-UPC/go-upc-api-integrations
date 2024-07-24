''
' Go UPC Microsoft Excel Integration Script
'
' @author      - Go-UPC
' @website     - https://go-upc.com
' @platform    - Mac OS
' @overview    - based on user input, this script automates the process of fetching
'                and filling in product data using Go-UPC's JSON API.
' @description - This script can be attached to any Excel spreadsheet using the
'                "VBA Macros" feature, which can be enabled in the application
'                Trust Center "Macro Settings" preferences.
''

'' Initialize and declare essentials
Option Explicit

Private Declare PtrSafe Function popen Lib "libc.dylib" (ByVal command As String, ByVal mode As String) As LongPtr
Private Declare PtrSafe Function pclose Lib "libc.dylib" (ByVal file As LongPtr) As Long
Private Declare PtrSafe Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
Private Declare PtrSafe Function feof Lib "libc.dylib" (ByVal file As LongPtr) As LongPtr

''
' Function execShell
'
' Executes a shell command in MS Excel
' [courtesy of Robert Knight via StackOverflow](stackoverflow.com/questions/6136798/vba-shell-function-in-office-2011-for-mac)
''
Function execShell(command As String, Optional ByRef exitCode As Long) As String
  Dim file As LongPtr
  file = popen(command, "r")

  If file = 0 Then
    Exit Function
  End If

  While feof(file) = 0
    Dim chunk As String
    Dim read As Long
    chunk = Space(50)
    read = fread(chunk, 1, Len(chunk) - 1, file)
    If read > 0 Then
      chunk = Left$(chunk, read)
      execShell = execShell & chunk
    End If
  Wend

  exitCode = pclose(file)
End Function

Function HTTPGet(sUrl As String, sQuery As String) As String

  Dim sCmd As String
  Dim sResult As String
  Dim lExitCode As Long

  sCmd = "curl --get -d """ & sQuery & """" & " " & """" & sUrl & """"
  sResult = execShell(sCmd, lExitCode)

  HTTPGet = sResult

End Function

''
' Public Function IndexOf:
'
' Searches for an item in a given collection and returns the index of the first occurrence of that item.
' If the item is found, the function returns its position in the collection; otherwise, it exits without returning a value.
''
Public Function IndexOf(ByVal coll As Collection, ByVal item As Variant) As Long
  Dim i As Long
  For i = 1 To coll.Count
    If coll(i) = item Then
      IndexOf = i
      Exit Function
    End If
  Next
End Function

''
' Function ClearRowValues:
'
' Clears the values in columns B through K (2 through 11) for a specified row in a specified sheet.
' This is useful for resetting the data in a row before updating it.
''
Function ClearRowValues(sheetNumber As Integer, rowNumber As Integer)
  Sheets(sheetNumber).Cells(rowNumber, 2).Value = ""
  Sheets(sheetNumber).Cells(rowNumber, 3).Value = ""
  Sheets(sheetNumber).Cells(rowNumber, 4).Value = ""
  Sheets(sheetNumber).Cells(rowNumber, 5).Value = ""
  Sheets(sheetNumber).Cells(rowNumber, 6).Value = ""
  Sheets(sheetNumber).Cells(rowNumber, 7).Value = ""
  Sheets(sheetNumber).Cells(rowNumber, 8).Value = ""
  Sheets(sheetNumber).Cells(rowNumber, 9).Value = ""
  Sheets(sheetNumber).Cells(rowNumber, 10).Value = ""
  Sheets(sheetNumber).Cells(rowNumber, 11).Value = ""
End Function

''
' Function SetRowDataNotFound:
'
' Clears the row values using ClearRowValues.
' Sets specific cells in the row to indicate that the product was not found, displaying the code type and a "(Product Not Found)" message.
''
Function SetRowDataNotFound(inputValue As String, codeType As String, sheetNumber As Integer, rowNumber As Integer)
  Call ClearRowValues(sheetNumber, rowNumber)
  Sheets(sheetNumber).Cells(rowNumber, 2).Value = codeType
  Sheets(sheetNumber).Cells(rowNumber, 3).Value = "(Product Not Found)"
End Function

''
' Private Sub Worksheet_Change:
'
' Triggers when any change occurs in the worksheet, 
' makes an API request, and fills in columns with 
' returned data accordingly.
''
Private Sub Worksheet_Change(ByVal Target As Range)
  Dim KeyCells As Range, SettingsCells As Range
  Dim http As Object, JSON As Object
  Dim i As Integer, sheetNum As Integer
  Dim apiKey As String, ProductCode As String, apiUrl As String
  Dim specsList As String, kv As String, codeType As String, gtinCode As String
  Dim specs, spec, prop

  '' Take first 1000 rows
  Set KeyCells = Range("A3:$A$1000")
  Set SettingsCells = Range("B1")

  apiKey = Range("UserAPIKey").Value

  If Not Application.Intersect(SettingsCells, Range(Target.Address)) Is Nothing Then
    MsgBox ("You can find your API Key on your account page: https://go-upc.com/account/profile")
    sheetNum = 1
    i = Target.Row
    If Len(Target.Value) >= 15 Then
      ' SettingsCells.EntireRow.Hidden = True
      MsgBox ("API key added. You may now start adding product codes (column A). Please only provide one code at a time.")
    End If
  '' The variable KeyCells contains the cells that trigger the automation sequence
  ElseIf Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
    If Len(apiKey) < 15 Then
      MsgBox ("Please provide your Go-UPC API key here (in column B1)")
    Else
      sheetNum = 1
      i = Target.Row

      If Target.Count > 1 Then
        MsgBox ("Please only update one code/row at a time!")
      Else
        If Len(Target.Value) >= 7 And Len(Target.Value) <= 14 Then
          '' Debug.Print "Value is Okay"
          ProductCode = Target.Value
          apiUrl = "https://go-upc.com/api/v1/code/" & ProductCode & "?key=" & apiKey
          Set JSON = ParseJson(HTTPGet(apiUrl, ""))

          Debug.Print JSON("codeType")

          Sheets(sheetNum).Cells(i, 1).Font.Color = RGB(0, 0, 0)

          If IsNull(JSON("product")) Then
            '' Set font color to gray to indicate no result
            Sheets(sheetNum).Cells(i, 1).Font.Color = RGB(20, 20, 20)
            If IsNull(JSON("codeType")) Then
              Call ClearRowValues(sheetNum, i)
            Else
              Call SetRowDataNotFound(ProductCode, JSON("codeType"), sheetNum, i)
            End If
          Else
            '' Extract Specs Data
            Set specs = JSON("product")("specs")
            specsList = ""
            kv = "key"
            For Each spec In specs
              If Len(specsList) > 0 Then
                specsList = specsList & ", "
              End If
              kv = "key"
              For Each prop In spec
                If kv = "key" Then
                  specsList = specsList & """" & prop & """: """
                  kv = "value"
                Else
                  specsList = specsList & prop & """"
                End If
              Next prop
            Next spec

            If Len(specsList) > 0 Then
              specsList = "{" & specsList & "}"
            End If

            codeType = Replace(LCase(JSON("codeType")), "-", "")

            if IsNull(JSON("product")(codeType)) Then
              gtinCode = ProductCode
            Else
              gtinCode = JSON("product")(codeType)
            End If

            Sheets(sheetNum).Cells(i, 2).Value = JSON("codeType")
            Sheets(sheetNum).Cells(i, 3).Value = JSON("product")("name")
            Sheets(sheetNum).Cells(i, 4).Value = JSON("product")("description")
            Sheets(sheetNum).Cells(i, 5).Value = JSON("product")("region")
            Sheets(sheetNum).Cells(i, 6).Value = JSON("product")("imageUrl")
            Sheets(sheetNum).Cells(i, 7).Value = JSON("product")("brand")
            Sheets(sheetNum).Cells(i, 8).Value = JSON("product")("category")
            Sheets(sheetNum).Cells(i, 9).Value = JSON("barcodeUrl")
            Sheets(sheetNum).Cells(i, 10).Value = gtinCode
            Sheets(sheetNum).Cells(i, 11).Value = specsList

          End If

        Else

          If Len(Target.Value) > 0 Then
            '' Set font color to red to indicate invalid code
            Sheets(sheetNum).Cells(i, 1).Font.Color = RGB(255, 0, 0)
            MsgBox ("Invalid UPC/EAN/ISBN!")
          End If

          Call ClearRowValues(sheetNum, i)
        End If
      End If
    End If
  End If
End Sub
