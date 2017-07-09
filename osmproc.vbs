Set objFSO=CreateObject("Scripting.FileSystemObject")

' ��� ����
outFile="log.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)

set xmlbody = createobject("Microsoft.XMLDOM")
xmlbody.Async="false"
xmlbody.load("input.osm")
xmlbody.setProperty "SelectionLanguage", "XPath"

' ����������� ������������� �������� �������������� ��������
Set MinusValues = xmlbody.selectNodes("/osm/node[@id<0]")
Set PlusValues  = xmlbody.selectNodes("/osm/node[@id>0]")
dim ArrPlusNODE()
i = 0
for each value in PlusValues
    If value.getAttribute("id") > i Then
    i = value.getAttribute("id")
    End If
next
' �������� i �� 1
i = i + 1

' ���������� ������������ ����� ������� 
' �������������� � ������ ��������������
' ���������� �������� NODE � ������ 
ReDim Preserve MyNODE1(1) ' ����� ������������� ������
ReDim Preserve MyNODE2(1) ' ������ ������������� ������
p = 0 ' ���������� node < 0

for each value in MinusValues
    MyNODE1(p) = i 
    MyNODE2(p) = CLng(value.getAttribute("id"))
    value.setAttribute "id", i
    objFile.WriteLine "NODE ID " & MyNODE1(p) & "   " & MyNODE2(p) 
    i = i + 1
    p = p + 1    
    ReDim Preserve MyNODE1(p+1)
    ReDim Preserve MyNODE2(p+1)
	value.setAttribute "version", 1
next
 
' ���������� ������������ ����� ������� 
' �������������� � ������ ��������������
' ���������� �������� WAY � ������ 
' ��������� ���������� � ���������� ID ��� NODE
Set NegData = xmlbody.selectNodes("/osm/way[@id<0]")
ReDim MyWAY1(1) '����� ������������� ������
ReDim MyWAY2(1) '������ ������������� ������
t = 0 ' ���������� way < 0
for each value in NegData
    MyWAY1(t) = i 
    MyWAY2(t) = CLng(value.getAttribute("id"))
    value.setAttribute "id", i
    objFile.WriteLine "WAY ID " & MyWAY1(t) & "   " & MyWAY2(t) 
    i = i + 1
    t = t + 1    
    ReDim Preserve MyWAY1(t+1)
    ReDim Preserve MyWAY2(t+1)
	value.setAttribute "version", 1
next

' �� ���� WAY, ������� ��������� �� NODE, ����� �������� 
' ������ ������������� �������� �� �����
Set NegData = xmlbody.selectNodes("/osm/way/nd[@ref<0]")
for each value in NegData
    tmp = CLng(value.getAttribute("ref"))
    objFile.WriteLine "WAY NEG NODE " & tmp
    for j = 0 to (p-1)   
        if tmp = MyNODE2(j) Then
            value.setAttribute "ref", MyNODE1(j)
        End If 
    next
next

' ���������� ������������ ����� ������� 
' �������������� � ������ ��������������
' ���������� �������� RELATION � ������ 
' ��������� ���������� � ���������� ID ��� WAY
Set NegData = xmlbody.selectNodes("/osm/relation[@id<0]")
ReDim MyREL1(1)
ReDim MyREL2(1)
m = 0 ' ���������� relation < 0
for each value in NegData
    MyREL1(m) = i 
    MyREL2(m) = CLng(value.getAttribute("id"))
    value.setAttribute "id", i
    objFile.WriteLine "REL ID " & MyREL1(m) & "   " & MyREL2(m) 
    i = i + 1
    m = m + 1    
    ReDim Preserve MyREL1(m+1)
    ReDim Preserve MyREL2(m+1)
	value.setAttribute "version", 1
next

' �� ���� RELATION, ������� ��������� �� NODE, ����� �������� 
' ������ ������������� �������� �� �����
Set NegData = xmlbody.selectNodes("/osm/relation/member[@ref<0 and @type='way']")
for each value in NegData
    tmp = CLng(value.getAttribute("ref"))
    objFile.WriteLine "WAY NEG REL " & tmp
    for j = 0 to (t-1)   
        if tmp = MyWAY2(j) Then
            value.setAttribute "ref", MyWAY1(j)
        End If 
    next
next

' �� ���� RELATION, ������� ��������� �� WAY, ����� �������� 
' ������ ������������� �������� �� �����
Set NegData = xmlbody.selectNodes("/osm/relation/member[@ref<0 and @type='node']")
for each value in NegData
    tmp = CLng(value.getAttribute("ref"))
    objFile.WriteLine "NODE NEG REL " & tmp
    for j = 0 to (p-1)   
        if tmp = MyNODE2(j) Then
            value.setAttribute "ref", MyNODE1(j)
        End If 
    next
next

' �� ���� RELATION, ������� ��������� �� RELATION, ����� �������� 
' ������ ������������� �������� �� �����
Set NegData = xmlbody.selectNodes("/osm/relation/member[@ref<0 and @type='relation']")
for each value in NegData
    tmp = CLng(value.getAttribute("ref"))
    objFile.WriteLine "RELATION NEG REL " & tmp
    for j = 0 to (m-1)   
        if tmp = MyREL2(j) Then
            value.setAttribute "ref", MyREL1(j)
        End If 
    next
next

xmlbody.save("output.osm")
objFile.Close