Dim ArrArrayOne: ArrArrayOne=Array(12,3,98,3,65,43,44,12,3,98,100,101)
Dim ArrArrayUnique()
Dim ObjUniqueDict: Set ObjUniqueDict=CreateObject("Scripting.Dictionary")
for iItr=0 to Ubound(ArrArrayOne)
    if not ObjUniqueDict.Exists(ArrArrayOne(iItr)) then
        ObjUniqueDict.Add ArrArrayOne(iItr),0
    end if
next
ReDim Preserve ArrArrayUnique(Ubound(ObjUniqueDict.keys()))
for kItr=0 to Ubound(ObjUniqueDict.keys())
    ArrArrayUnique(kItr)=ObjUniqueDict.keys()(kItr)
next
msgbox(join(ArrArrayUnique,"-"))