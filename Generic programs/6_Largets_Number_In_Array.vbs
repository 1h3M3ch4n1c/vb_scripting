Dim IntArrNumbersArray:IntArrNumbersArray=Array(0, 123, 46,987,34762,7623,62,101)
iTemp=0
for iItr=0 to Ubound(IntArrNumbersArray)
    if IntArrNumbersArray(iItr)>iTemp Then
        iTemp=IntArrNumbersArray(iItr)
    end if
next
msgbox(iTemp)