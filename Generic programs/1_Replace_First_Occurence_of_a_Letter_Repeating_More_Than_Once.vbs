StrText="Uniqe Uniqe Uniqe"
iLengthStrText=len(StrText)
strUNique=""
Set objDictStrText=CreateObject("Scripting.Dictionary")
for iItr=0 to iLengthStrText
    objDictStrText(iItr)=Mid(StrText,iItr+1,1)
next

for iItr=0 to iLengthStrText
    iCount=0
    ChrICharacter=Mid(StrText,iItr+1,1)
    
    for jItr=0 to iLengthStrText
        ChrJCharacter=Mid(StrText,jItr+1,1)
        if (strcomp(" ",ChrICharacter)<>0 and strcomp(lcase(ChrJCharacter),lcase(ChrICharacter))=0) Then
            iCount=iCount+1   
        end if
    next
    if iCount>=2 Then
        for kItr=0 to iLengthStrText
            if (strcomp(lcase(objDictStrText(kItr)),lcase(ChrICharacter))=0  and InStr(1,strUNique,lcase(objDictStrText(kItr)))=0) Then
                iFound=kItr
                strUNique=strUNique&lcase(objDictStrText(kItr))
                Exit For
            end if
        next
        objDictStrText(iFound)=0
    end if
next
msgbox(join(objDictStrText.items(),""))