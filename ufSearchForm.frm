VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSearchForm 
   Caption         =   "Inventory Information"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5880
   OleObjectBlob   =   "ufSearchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSearchForm"
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

Private Sub Clear_Click()
'Reloads form fresh

Worksheets("Variables").Range("A1").Clear
Unload Me
ufSearchForm.Show

End Sub

Private Sub LogUse_Click()
'Unhides consumption fields and save button.

LogUse.Visible = False
lbUsage.Visible = True
txtUsed.Visible = True
Save.Visible = True

End Sub

Private Sub Save_Click()
'Logs consumption data into Consumption table.

Dim tbl As ListObject
Dim newrow As ListRow
Dim ws As Worksheet

Set ws = Worksheets("Variables")
Set tbl = Worksheets("Tables").ListObjects("Consumption")
Set newrow = tbl.ListRows.Add

With newrow            'Logs consumption.
    .Range(1) = Now()
    .Range(2) = Me.txtInventoryCode.Value
    .Range(3) = Me.txtUsed.Value
End With

With Me                 'Updates display with new quantity onhand.
    .lbBrand.Caption = ws.Range("A2")
    .lbMaterial.Caption = ws.Range("A3")
    .lbColor.Caption = ws.Range("A4")
    .lbTemps.Caption = ws.Range("A5")
    .lbQOH.Caption = ws.Range("A6")
    .txtUsed.Value = ""
End With

ThisWorkbook.Save
MsgBox "Consumption logged. New volume on hand is " & ws.Range("A6") & ".", vbOKOnly

'Disables consumption fields
lbUsage.Visible = False
txtUsed.Visible = False
Save.Enabled = False
Save.Visible = False
LogUse.Visible = True

End Sub

Private Sub Search_Click()
'Enables Use Material button, enters input Inventory Code into temporary variable, posts returned results onto form

LogUse.Enabled = True

Dim ws As Worksheet
Set ws = Worksheets("Variables")

ws.Range("A1").Value = Me.txtInventoryCode.Value

With Me
    .lbBrand.Caption = ws.Range("A2")
    .lbMaterial.Caption = ws.Range("A3")
    .lbColor.Caption = ws.Range("A4")
    .lbTemps.Caption = ws.Range("A5")
    .lbQOH.Caption = ws.Range("A6")
    .txtUsed.Value = ""
End With


End Sub

Private Sub txtInventoryCode_Change()
'Enables Clear option

Clear.Enabled = True

End Sub

Private Sub UserForm_Terminate()
'Clears temporary variables, saves workbook.

Worksheets("Variables").Range("A1").Clear
ThisWorkbook.Save

End Sub
