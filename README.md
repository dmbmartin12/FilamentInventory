# FilamentInventory
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

A small database built in Excel with some macros to make managing 3D printing filament a breeze. 

After starting up a 3D printing project that will have me using several colors, it occured to me that I don't have a good way to track how much filament I have left on a spool.

With that in mind, I started out with a simple table, then added a little more and a little more until eventually I came up with this little gem and thought maybe someone else could make use of it. 

I included every userform with complete layouts of the macros I built for you to review prior to running. That said - remember that it is risky to run any scripts or macros from unknown sources, even me. This is offered without warranty and As-Is using MIT license. You may edit and redistribute for non-commercial purposes as long as you keep my copywright statement. 

I plan to add some things like links for reordering and a reporting function that shows the history of an item. 

How to use:
To add new inventory:
Step 1. Open the Excel file. Give affirmative answers to security warnings. 
Step 2. Choose "Log New Inventory" from the menu.
Step 3. Fill in the blanks.
  For Brand, Material, Color, and Vendor, press "New" to add a new option to the list.
Step 4. Click Save. You will receive a pop-up message that confirms the item has been added to your inventory and what the ID number is. 
  Write this ID number on the spool for future reference. Personally, I have a Dymo printer and was able to print labels. A sharpee will do, too.

To view the inventory status:
Step 1. Open excel file. Give affirmative answers to security warnings.
Step 2. Choose "Consume Current Stock"
Step 3. Enter the Inventory ID Code and click Search.
  Current supply level will be displayed. 

To log consumption:
Step 1. Follow steps for Viewing Inventory Status above.
Step 2. Click "Use Material"
Step 3. Enter the amount of filament your print is expected to consume. Your slicer should give you this figure.
Step 4. Click Save. You will be presented with a confrimation letting you know your consumption has been logged. 

If you want to have a look at the tables or to open VB, press the Debug button from the Welcome menu. Take care not to make changes to the tables in the "Tables" worksheet or anything in column A of "Variables".

Please leave me your feedback. If you have a feature request, let me know and I'll see what I can do. 
