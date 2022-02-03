VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufNewColor 
   Caption         =   "New Color"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3000
   OleObjectBlob   =   "ufNewColor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufNewColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MIT License
'Copyright (c) 2022 Tyler Jones, tyler.jones@mymail.centralpenn.edu
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,'
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Private Sub Submit_Click()
'Adds input Color onto variable table and new Color table.

Dim ws As Worksheet
Dim ws2 As Worksheet
Dim tbl As ListObject
Dim newrow As ListRow

Set ws = Worksheets("Tables")
Set ws2 = Worksheets("Variables")
Set tbl = Worksheets("Tables").ListObjects("ColorTable")
Set newrow = tbl.ListRows.Add

With newrow
    .Range(1) = Me.ColorInput.Value
    .Range(2) = ws2.Range("A15")
    
End With

ws2.Range("A9") = Me.ColorInput.Value

Unload Me
ufNewStock.Show

End Sub

Private Sub UserForm_Initialize()
'Enters the typed color into input field.

Me.ColorInput.Value = Worksheets("Variables").Range("A9")

End Sub

Private Sub UserForm_Terminate()
'Opens New Stock userform

ufNewStock.Show
End Sub
