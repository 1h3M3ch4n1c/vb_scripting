Dim ArrIntNumbersArray:ArrIntNumbersArray=Array(12,32,67,9,2,31,456)
iCount=0
for iItr=0 to Ubound(ArrIntNumbersArray)
    iCount=iCount+ArrIntNumbersArray(iItr)
next
msgbox(iCount)