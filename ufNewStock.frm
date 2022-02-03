VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufNewStock 
   Caption         =   "New Inventory"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6930
   OleObjectBlob   =   "ufNewStock.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufNewStock"
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
'Clears the variables saved, closes and relaunches this form

Worksheets("Variables").Range("A7:A13").Clear

Unload Me
ufNewStock.Show

End Sub

Private Sub NewBrand_Click()
'Sets temporary variables, opens "New Brand" dialog box

Dim ws As Worksheet
Set ws = Worksheets("Variables")

With ws
    .Range("A7") = Me.cbBrand
    .Range("A8") = Me.cbMaterials
    .Range("A9") = Me.cbColors
    .Range("A10") = Me.Temp_TXT
    .Range("A11") = Me.Volume_txt
    .Range("A12") = Me.Cost_txt
    .Range("A13") = Me.cbVendors
End With

Unload Me
ufNewBrand.Show

End Sub

Private Sub NewColor_Click()
'Sets temporary variables, opens "New Color" dialog box

Dim ws As Worksheet
Set ws = Worksheets("Variables")

With ws
    .Range("A7") = Me.cbBrand
    .Range("A8") = Me.cbMaterials
    .Range("A9") = Me.cbColors
    .Range("A10") = Me.Temp_TXT
    .Range("A11") = Me.Volume_txt
    .Range("A12") = Me.Cost_txt
    .Range("A13") = Me.cbVendors
End With

Unload Me
ufNewColor.Show
End Sub

Private Sub NewMaterial_Click()
'Sets temporary variables, opens "New Material" dialog box

Dim ws As Worksheet
Set ws = Worksheets("Variables")


With ws
    .Range("A7") = Me.cbBrand
    .Range("A8") = Me.cbMaterials
    .Range("A9") = Me.cbColors
    .Range("A10") = Me.Temp_TXT
    .Range("A11") = Me.Volume_txt
    .Range("A12") = Me.Cost_txt
    .Range("A13") = Me.cbVendors
End With

Unload Me
ufNewMaterial.Show
End Sub

Private Sub NewVendor_Click()
'Sets temporary variables, opens "New Vendor" dialog box

Dim ws As Worksheet
Set ws = Worksheets("Variables")


With ws
    .Range("A7") = Me.cbBrand
    .Range("A8") = Me.cbMaterials
    .Range("A9") = Me.cbColors
    .Range("A10") = Me.Temp_TXT
    .Range("A11") = Me.Volume_txt
    .Range("A12") = Me.Cost_txt
    .Range("A13") = Me.cbVendors
End With

Unload Me
ufNewVendor.Show
End Sub

Private Sub Save_Click()
'Saves all temporary vairables into the Master table, displays confrimation screen, clears temporary files, opens Welcome dialog.

Dim ws As Worksheet
Dim ws2 As Worksheet
Dim tbl As ListObject
Dim newrow As ListRow

Set ws = Worksheets("Tables")
Set ws2 = Worksheets("Variables")
Set tbl = ws.ListObjects("Master")
Set newrow = tbl.ListRows.Add

'Sets current input to temporary variables
With ws2
    .Range("A7") = Me.cbBrand
    .Range("A8") = Me.cbMaterials
    .Range("A9") = Me.cbColors
    .Range("A10") = Me.Temp_TXT
    .Range("A11") = Me.Volume_txt
    .Range("A12") = Me.Cost_txt
    .Range("A13") = Me.cbVendors
End With


'Saves current most recent Inventory Code as a variable to be checked against later.
With ws2
    .Range("A19").Copy
    .Range("A14").PasteSpecial Paste:=xlPasteValues
End With


'Copies temporary data into Master table
With newrow
    .Range(1) = ws2.Range("A19") + 1
    .Range(2) = ws2.Range("A8")     'Material
    .Range(3) = ws2.Range("A9")     'Color
    .Range(4) = ws2.Range("A11")    'Starting Volume
    .Range(5) = ws2.Range("A10")    'Suggested printing temps
    .Range(6) = ws2.Range("A13")    'Vendor
    .Range(7) = ws2.Range("A7")     'Brand
    .Range(8) = ws2.Range("A12")    'Price
    .Range(9) = ws2.Range("A20")    'Purchase date (default to today)
    .Range(10) = Now()              'Time Stamp
End With

'Checks to see if save was successful by comparing previous most recent inventory Code with the current most recent.
If ws2.Range("A19") <> ws.Range("A14") Then
    
        With ws2
            .Range("A7:A14").Clear
            End With
        MsgBox "Success! The ID number for this asset is " & ws2.Range("A19") & ".", vbOKOnly
        Unload Me
        ufWelcomeSplash.Show
    Else
        MsgBox "Something went wrong. Please try again.", vbOKOnly
End If

End Sub

Private Sub UserForm_Initialize()
'Prefills combo boxes with all possible choices, displays form prefilled with entries stored temporary variable fields.

Dim ws As Worksheet
Dim ws2 As Worksheet
Dim cLoc As Range

Set ws = Worksheets("Tables")
Set ws2 = Worksheets("Variables")


For Each cLoc In ws.Range("Brands")     'Populates cbBrands
    With Me.cbBrand
    .AddItem cLoc.Value
    End With

Next cLoc

For Each cLoc In ws.Range("Materials")  'Populates cbMaterials
    With Me.cbMaterials
    .AddItem cLoc.Value
    End With

Next cLoc

For Each cLoc In ws.Range("Colors")     'Populates Colors
    With Me.cbColors
    .AddItem cLoc.Value
    End With

Next cLoc

For Each cLoc In ws.Range("Vendors")    'Populates Vendors
    With Me.cbVendors
    .AddItem cLoc.Value
    End With
Next cLoc

With Me
    .cbBrand = ws2.Range("A7")          'Populates all other fields
    .cbMaterials = ws2.Range("A8")
    .cbColors = ws2.Range("A9")
    .Temp_TXT = ws2.Range("A10")
    .Volume_txt = ws2.Range("A11")
    .Cost_txt = ws2.Range("A12")
    .cbVendors = ws2.Range("A13")
End With

End Sub

Private Sub UserForm_Terminate()
'Saves the workbook

ThisWorkbook.Save

End Sub
