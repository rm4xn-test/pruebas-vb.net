Imports System.Text.RegularExpressions

Friend Module Module1
  
  Private Alphabet As Char()
  Private LastColumn As String
  
  Public Sub Main()
    Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
    LastColumn = "AAA"
    
    Console.WriteLine(Numbers("B16"))
    ExcelColumns()
    Console.Read()
  End Sub
  
  Public Function Numbers(text As String) As Integer
    Return Regex.Replace(text, "[^0-9]", "")
  End Function
  
  Public Sub ExcelColumns()
    Dim columns As New List(Of String)
    For i As Integer = 1 To LastColumn.Length
      If Not LoadColumns("", i) Then Exit For
    Next
  End Sub
  
  Private Sub LoadColumns(start as String, counter as Integer) As Boolean
    For i As Integer = 0 To Alphabet.Length - 1
      Dim temp As String = start
      temp += Alphabet(i)
      If temp.Length = counter Then
        Columns.Add(temp)
      ElseIf Not LoadColumns(temp, counter) Then
        Return False
      End If
      If temp = LastColumn Then Return False
    Next
    Return True
  End Sub
