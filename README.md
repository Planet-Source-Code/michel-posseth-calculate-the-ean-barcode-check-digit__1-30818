<div align="center">

## calculate the EAN \( barcode\) check digit


</div>

### Description

ever made a program capable of showing barcodes ?

if you did than you`ve been there ,,, at the oficial EAN standards site,, than you would have seen how to calculate the check digit.

they hold the standard they publish that standard also on their website ,,,

http://www.ean-int.org/index800.html

i never found code in VB that calculates the check digit ,, so my conclusion is that it was hold for comercial reassons ( there are lots of controls out there for a lot of monney :-)

so i donate this M. Posseth code to the public and make it public domain ,,,

uhmmm votes would be apreciated :-)
 
### More Info
 
13 digit EAN code ( manufacterer number , parts number )

14 digit EAN code ( calculates the check digit )


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michel Posseth](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michel-posseth.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michel-posseth-calculate-the-ean-barcode-check-digit__1-30818/archive/master.zip)





### Source Code

```
Private Function CalcEanMetcontrole(ByVal EAN13Digit As String) As String
Dim Explodestring As String
Dim DigArray
Dim Digit As Variant
Dim factor As Integer
Dim Standin As Integer
Dim som As Integer
Dim CG As Integer
Explodestring = Left$(Replace(StrConv(EAN13Digit, vbUnicode), vbNullChar, _
        ","), Len(EAN13Digit) * 2 - 1)
  DigArray = Split(Explodestring, ",", -1, 1)
factor = 3
For Each Digit In DigArray
Standin = CInt(Digit)
som = som + (Standin * factor)
factor = 4 - factor
Next
If Right$(CStr(som), 1) = 0 Then
CG = 0
Else
CG = 10 - Right$(som, 1)
End If
CalcEanMetcontrole = Trim$(EAN13Digit & CStr(CG))
End Function
```

