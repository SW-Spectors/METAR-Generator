Option Explicit
'Const CREATE_FILE

Dim MSG, a, b, c, d, di, de, ei, e, f, g, h1, h2, h, i, j, k, l, m, n, o
Dim fso, data

a = InputBox("�n�旪�������", "METAR_Generator")
b = InputBox("�ϑ����������(UTC)[��+����]", "METAR_Generator") & "Z"
c = InputBox("���������(�x)", "METAR_Generator")
di = InputBox("���������(m/s)", "METAR_Generator")
de = di * 1.943844
d = Round(de) & "KT"
e = InputBox("���������(m)", "METAR_Generator")
'if ei >= 10000 Then
'	e = 9999
'if ei >= 5000 and ei < 10000 Then
'	e = Round(ei, -3)
'if ei < 5000 Then
'	e = Round(ei, -2)
f = InputBox("���݂̓V�C�����", "METAR_Generator")
g = InputBox("�_�ʁE�_��̍��������", "METAR_Generator")
h1 = InputBox("�C�������", "METAR_Generator")
h2 = InputBox("�I�_�����", "METAR_Generator")
h = h1 & "/" & h2
j = "Q" & InputBox("�C�������[hPa]", "METAR_Generator")
k = InputBox("�����w����?[������΂��̐悩��͋󔒂�OK]", "METAR_Generator")
l = InputBox("�_�ʁE�_�`�E�_��̍���", "METAR_Generator")
m = "A" & InputBox("�C�������[inHg]", "METAR_Generator")
MSG = "METAR " & a & " " & b & " " & c & d & " " & e & " " & f & " " & g & " " & h & " " & j & " " & k & " " & l & " " & m
o = MsgBox("METAR " & a & " " & b & " " & c & d & " " & e & " " & f & " " & g & " " & h & " " & j & " " & k & " " & l & " " & m, vbOKOnly + vbExclamation + vbApplicationModal + vbSystemModal, "Generated")

Set fso = CreateObject("Scripting.FileSystemObject")
Set data = fso.CreateTextFile("METAR_GeneratedFile.txt", true)
data.WriteLine(MSG)
data.Close
Set fso = Nothing