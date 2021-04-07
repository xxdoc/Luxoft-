VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim words() As String
Dim minus As Long
Dim counter As Long
Dim arrays As String

'Main function
Public Function Pairs(a As String, b As String, c As String) As String

'We assigning the array to the variable
arrays = c

'Stating the counter in 1
counter = 1

'Convert the string to Array with the function Split()
words() = Split(arrays)

'Stating the variable in order to use in the For Each loop
Dim Element

'Using a For loop and For Each loop in order to validate the values
For i = LBound(words) To UBound(words)
        For Each Element In words
        minus = CInt(i) - CInt(Element)
        'If (minus = CInt(numcomparative)) Then
        If (minus = CInt(b)) Then
        counter = counter + 1
        End If
        Next
Next i

'Return the value as a message
MsgBox (counter)

End Function

