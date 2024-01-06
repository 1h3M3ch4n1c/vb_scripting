StrText="This is a string"
ILengthStrText=len(StrText)
StrTextReverse=""
for iItr=0 to ILengthStrText-1
    StrTextReverse=Mid(StrText,iItr+1,1)&StrTextReverse
next
msgbox(StrTextReverse)