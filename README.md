<div align="center">

## AmtToWords


</div>

### Description

This function converts amount in words with supplied currency parameters.

e.g. AmtToWords(12345.01, "GB POUND", "PENNY", "GB POUNDS", "PENNIES") will return GB POUNDS TWELVE THOUSAND THREE HUNDRED FORTY-FIVE and ONE PENNY ONLY
 
### More Info
 
Amount As Currency, UnitCurr As String, DecCurr As String, UnitsCurr As String, DecsCurr As String

max amount that can be converted by this function is 922,337,203,685,477

It returns amount in words with Currency Parameters


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Atul Alurkar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/atul-alurkar.md)
**Level**          |Unknown
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/atul-alurkar-amttowords__1-1083/archive/master.zip)

### API Declarations

None


### Source Code

```
Private units(20), teens(11)
Function AmtToWords(amount As Currency, UnitCurr As String, DecCurr As String, UnitsCurr As String, DecsCurr As String) As String
Dim new_amt, TRstring, BIstring, MIstring, THstring, HUstring, DEstring, Separator As String
If amount = 0 Then
  AmtToWords = "NIL"
  Exit Function
End If
units(0) = ""
units(1) = " ONE"
units(2) = " TWO"
units(3) = " THREE"
units(4) = " FOUR"
units(5) = " FIVE"
units(6) = " SIX"
units(7) = " SEVEN"
units(8) = " EIGHT"
units(9) = " NINE"
units(10) = " TEN"
units(11) = " ELEVEN"
units(12) = " TWELVE"
units(13) = " THIRTEEN"
units(14) = " FOURTEEN"
units(15) = " FIFTEEN"
units(16) = " SIXTEEN"
units(17) = " SEVENTEEN"
units(18) = " EIGHTEEN"
units(19) = " NINETEEN"
teens(0) = ""
teens(1) = " TEN"
teens(2) = " TWENTY"
teens(3) = " THIRTY"
teens(4) = " FORTY"
teens(5) = " FIFTY"
teens(6) = " SIXTY"
teens(7) = " SEVENTY"
teens(8) = " EIGHTY"
teens(9) = " NINETY"
teens(10) = " HUNDRED"
new_amt = Format(amount, "000000000000000.00")
TRstring = Mid(new_amt, 1, 3)
BIstring = Mid(new_amt, 4, 3)
MIstring = Mid(new_amt, 7, 3)
THstring = Mid(new_amt, 10, 3)
HUstring = Mid(new_amt, 13, 3)
DEstring = "0" + Mid(new_amt, 17, 2)
AmtToWords = ""
UnitCurr = IIf(Val(Left(new_amt, 15)) = 0, "", UnitCurr)
DecCurr = IIf(Val(Right(new_amt, 2)) = 0, "", DecCurr)
UnitCurr = IIf(Val(Left(new_amt, 15)) > 1, UnitsCurr, UnitCurr)
DecCurr = IIf(Val(Right(new_amt, 2)) > 1, DecsCurr, DecCurr)
Separator = IIf(UnitCurr <> "" And DecCurr <> "", " and", "")
AmtToWords = UnitCurr + AmtToWords
AmtToWords = AmtToWords + IIf(Val(TRstring) > 0, numconv(TRstring) + " TRILLION", "")
AmtToWords = AmtToWords + IIf(Val(BIstring) > 0, numconv(BIstring) + " BILLION", "")
AmtToWords = AmtToWords + IIf(Val(MIstring) > 0, numconv(MIstring) + " MILLION", "")
AmtToWords = AmtToWords + IIf(Val(THstring) > 0, numconv(THstring) + " THOUSAND", "")
AmtToWords = AmtToWords + IIf(Val(HUstring) > 0, numconv(HUstring), "")
AmtToWords = AmtToWords + IIf(Val(DEstring) > 0, Separator + numconv(DEstring), "")
AmtToWords = Trim(AmtToWords + " " + DecCurr) + " ONLY"
End Function
Function numconv(amt) As String
Dim aAmount, bAmount, cAmount, dAmount As Integer
Dim hyphen As String
aAmount = Val(Mid(amt, 2, 2))
bAmount = Val(Mid(amt, 3, 1))
cAmount = Val(Mid(amt, 2, 1))
dAmount = Val(Mid(amt, 1, 1))
If aAmount < 20 Then
  numconv = units(aAmount)
Else
  numconv = units(bAmount)
  If bAmount > 0 And cAmount > 0 Then
    hyphen = "-"
  End If
  numconv = teens(cAmount) + hyphen + LTrim(numconv)
End If
If dAmount > 0 Then
  numconv = units(dAmount) + " HUNDRED" + numconv
End If
End Function
```

