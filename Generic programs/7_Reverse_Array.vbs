Dim IntArrNumbersArray:IntArrNumbersArray=Array(123,34,6,90,34,21,6,789)
Dim IntArrNumbersArrayReversed:Redim IntArrNumbersArrayReversed(Ubound(IntArrNumbersArray))
iLengthOriginalArray=Ubound(IntArrNumbersArray)
' msgbox(iLengthOriginalArray)
for iItr=Ubound(IntArrNumbersArray) to 0 Step -1
    IntArrNumbersArrayReversed(iLengthOriginalArray-iItr)=IntArrNumbersArray(iItr)
next
msgbox(join(IntArrNumbersArrayReversed,"|"))