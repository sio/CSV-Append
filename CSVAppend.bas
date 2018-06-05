Attribute VB_Name = "CSVAppend"
Option Explicit

'
' Append to CSV from Excel without breaking its structure
'
' Editor UI is a worksheet (named Append by default) that contains the following
' named ranges:
'   - FilePath
'       Full path to the CSV file. Relative paths are supported
'   - Delimiter
'       Field delimiter. Defaults to semicolon (;)
'   - Quote
'       Quote character for escaping values containing delimiter. Defaults to
'       double quote (")
'   - Charset
'       Text file encoding. See documentation for Adodb.Stream for the list of
'       supported values. Defaults to "utf-8"
'   - EOL
'       Line breaks type. Either "CRLF" (dos) or "LF" (unix)
'   - DataArea
'       Describes the area that's safe for editing. Only top left cell matters
'


'
' Copyright 2018 Vitaly Potyarkin
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'


Public Sub OpenCSV(Optional CurrentSheet As Worksheet)
    If CurrentSheet Is Nothing Then
        Set CurrentSheet = Append
    End If

    Dim Settings As Collection
    Set Settings = GetSettings(CurrentSheet)

    Dim Header() As String
    Header = GetCSVHeader(Settings)

    Dim i
    Dim Cell
    Dim ThisRange As Range
    Dim Exclude As String
    On Error Resume Next
    For Each i In Array("NewRow", "NewRowHeader")
        Set ThisRange = CurrentSheet.Range(i)
        ThisRange.ClearContents
        Exclude = ThisRange.Cells(1, 1).Address
        For Each Cell In ThisRange
            If Cell.Cells(1, 1).Address <> Exclude Then
                Cell.Clear
            End If
        Next
    Next
    On Error GoTo 0

    Dim NewRow As Range
    Dim NewRowHeader As Range
    For i = LBound(Header) To UBound(Header)
        Set Cell = Settings("DataArea").Cells(1, 1).Offset(i + 1, 0)
        Cell.Value = Header(i)
        If NewRow Is Nothing Then
            Set NewRow = Cell.Offset(0, 1)
            Set NewRowHeader = Cell
        Else
            Set NewRow = Union(NewRow, Cell.Offset(0, 1))
            Set NewRowHeader = Union(NewRowHeader, Cell)
        End If
    Next

    CurrentSheet.Parent.Names.Add Name:="NewRow", RefersTo:=NewRow
    CurrentSheet.Parent.Names.Add Name:="NewRowHeader", RefersTo:=NewRowHeader

    For Each i In Array("NewRowHeader", "NewRow")
        With CurrentSheet.Range(i)
            .Cells(1, 1).Copy
            .PasteSpecial xlPasteFormats
        End With
    Next
    Application.CutCopyMode = False
End Sub


Public Sub WriteCSV(Optional CurrentSheet As Worksheet)
    If CurrentSheet Is Nothing Then
        Set CurrentSheet = Append
    End If

    Dim Settings As Collection
    Set Settings = GetSettings(CurrentSheet)

    Dim NextLine As String
    Const InitialValue As String = "__Initial_value__;"","
    NextLine = InitialValue

    Dim Value As String
    Dim Cell
    For Each Cell In CurrentSheet.Range("NewRow")
        Value = CStr(Cell.Value)
        If InStr(Value, Settings("Delimiter")) <> 0 Then
            Value = Settings("Quote") + Value + Settings("Quote")
        End If
        If NextLine = InitialValue Then
            NextLine = Value
        Else
            NextLine = NextLine + Settings("Delimiter") + Value
        End If
    Next

    With CreateObject("ADODB.Stream")
        .Charset = Settings("Charset")
        .Type = 2  ' Text data
        .Open
        .LoadFromFile (Settings("FilePath"))
        .Position = .Size - Len(Settings("EOL"))
        If .ReadText(-1) <> Settings("EOL") Then
            NextLine = Settings("EOL") + NextLine
        End If
        .WriteText NextLine + Settings("EOL")
        .SaveToFile Settings("FilePath"), 2  ' adSaveCreateOverWrite
        .Close
    End With

    CurrentSheet.Range("NewRow").ClearContents
End Sub


Private Function GetSettings(Sh As Worksheet) As Collection
    Dim DefaultSettings As New Collection
    DefaultSettings.Add ";", "Delimiter"
    DefaultSettings.Add """", "Quote"
    DefaultSettings.Add "CRLF", "EOL"
    DefaultSettings.Add "utf-8", "Charset"

    Dim Settings As New Collection
    Dim Key, Value
    For Each Key In Array("FilePath", "Delimiter", "Quote", "EOL", "DataArea", "Charset")
        ' Do not fail if named range is not found
        On Error Resume Next
        Value = ""
        Value = Sh.Range(Key).Value
        On Error GoTo 0

        ' Fall back to defaults
        If Value = "" Or Value = 0 Then
            Value = DefaultSettings(Key)
        End If

        ' Some special cases below:
        ' - Transform EOL name to actual character
        If Key = "EOL" Then
            If Value = "CRLF" Then
                Value = vbCrLf
            ElseIf Value = "LF" Then
                Value = vbLf
            Else
                Err.Raise 1000, "GetSettings", "Invalid line feed value"
            End If
        End If
        ' - Store DataArea as a reference to the range
        If Key = "DataArea" Then
            Set Value = Sh.Range(Key)
        End If
        ' - Support relative FilePath
        ChDrive Left(Sh.Parent.Path, InStr(Sh.Parent.Path, ":") - 1)
        ChDir Sh.Parent.Path

        Settings.Add Value, Key
    Next

    Set GetSettings = Settings
End Function


Private Function GetCSVHeader(Settings As Collection)
    Dim RawHeader As String
    Dim Header() As String
    Dim TempHeader As New Collection
    Dim FieldName As String
    Dim Char As String
    Dim Position As Long
    Dim NumberOfFields As Long
    Dim QuoteOpen As Boolean

    With CreateObject("ADODB.Stream")
        .Charset = Settings("Charset")
        .Open
        .LoadFromFile (Settings("FilePath"))
        If Settings("EOL") = vbCrLf Then
            .LineSeparator = -1
        ElseIf Settings("EOL") = vbLf Then
            .LineSeparator = 10
        ElseIf Settings("EOL") = vbCr Then
            .LineSeparator = 13
        Else
            Err.Raise 1000, "Invalid line separator"
        End If
        RawHeader = .ReadText(-2)  ' Read one line
        .Close
    End With

    NumberOfFields = 0
    FieldName = ""
    QuoteOpen = False
    For Position = 1 To Len(RawHeader)
        Char = Mid(RawHeader, Position, 1)
        If Char = Settings("Quote") Then
            QuoteOpen = Not QuoteOpen
        ElseIf Char = Settings("Delimiter") And Not QuoteOpen Then
            TempHeader.Add FieldName
            NumberOfFields = NumberOfFields + 1
            FieldName = ""
        Else
            FieldName = FieldName + Char
        End If
    Next
    TempHeader.Add FieldName
    NumberOfFields = NumberOfFields + 1

    ReDim Header(NumberOfFields - 1)
    For Position = LBound(Header) To UBound(Header)
        Header(Position) = TempHeader(Position + 1)
    Next
    Set TempHeader = Nothing
    GetCSVHeader = Header
End Function
