Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports System.Console
Module Module1
    Public Sub Main()
        Dim P As String
        Dim N As String
        Dim obj As Object
        WriteLine("Please input path for CSV file")
        P = FormatPath(ReadLine())
        WriteLine("Please input Name for CSV file")
        N = ReadLine()
        If Not P = vbNullString And Not N = vbNullString Then
            obj = MainFun(P, N)
        End If
    End Sub
    Function MainFun(path As String, nameF As String) As List(Of String())
        If File.Exists(path & nameF) = False Then Return Nothing
        Return VBCSVParser(path & nameF)
    End Function
    Public Function VBCSVParser(ByVal path As String) As List(Of String())
        'Parser CSV jeszcze z czasow VB (6.0) nadal najlepsze z wszystkiego co C# moze w tej
        ' kwesti zaoferowac -Dariusz Tomczak

        Dim tempList As List(Of String()) = New List(Of String())()

        Dim csvParser As TextFieldParser = New TextFieldParser(path)
        With csvParser
            .CommentTokens = New String() {"#"}
            .SetDelimiters(New String() {";"})
            .HasFieldsEnclosedInQuotes = True
            .ReadLine()
        End With
        While Not csvParser.EndOfData
            ' Read current line fields, pointer moves to the next line.
            Dim fields As String() = csvParser.ReadFields()
            tempList.Add(fields)
        End While

        tempList = tempList.OrderBy(Function(arr) arr(0)).ThenBy(Function(arr) arr(1)).ToList() 'Linq sort
        Return tempList
    End Function
    Public Function FormatPath(ByVal path As String) As String
        Dim slash = "\" ' check if path is ending with slash
        Dim c As Char = path.Trim().ToString().Last()
        If Not Equals(c.ToString(), slash) Then path += slash
        Return path
    End Function
End Module
