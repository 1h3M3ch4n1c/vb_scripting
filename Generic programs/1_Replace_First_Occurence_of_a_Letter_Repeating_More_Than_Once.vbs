StrText="This is a new string"
iLengthStrText=len(StrText)
strUNique=""
for iItr=1 to iLengthStrText
    iCount=0
    ChrICharacter=Mid(StrText,iItr,1)
    strUNique=strUNique&ChrICharacter
    for jItr=1 to iLengthStrText
        ChrJCharacter=Mid(StrText,jItr,1)
        if (strcomp(ChrJCharacter,ChrICharacter)=0 and InStr(1,strUnique,ChrJCharacter)<>0) Then
            iCount=iCount+1   
        end if
    next
    if iCount>2 Then
        strReplaced=Replace(StrText,ChrICharacter,0)
    end if
next
msgbox(strReplaced)
