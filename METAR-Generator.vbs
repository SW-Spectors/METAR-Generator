Option Explicit
'Const CREATE_FILE

Dim MSG, a, b, c, d, di, de, ei, e, f, g, h1, h2, h, i, j, k, l, m, n, o
Dim fso, data

a = InputBox("地域略号を入力", "METAR_Generator")
b = InputBox("観測日時を入力(UTC)[日+時間]", "METAR_Generator") & "Z"
c = InputBox("風向を入力(度)", "METAR_Generator")
di = InputBox("風速を入力(m/s)", "METAR_Generator")
de = di * 1.943844
d = Round(de) & "KT"
e = InputBox("視程を入力(m)", "METAR_Generator")
'if ei >= 10000 Then
'	e = 9999
'if ei >= 5000 and ei < 10000 Then
'	e = Round(ei, -3)
'if ei < 5000 Then
'	e = Round(ei, -2)
f = InputBox("現在の天気を入力", "METAR_Generator")
g = InputBox("雲量・雲梯の高さを入力", "METAR_Generator")
h1 = InputBox("気温を入力", "METAR_Generator")
h2 = InputBox("露点を入力", "METAR_Generator")
h = h1 & "/" & h2
j = "Q" & InputBox("気圧を入力[hPa]", "METAR_Generator")
k = InputBox("国内指示符?[無ければこの先からは空白でOK]", "METAR_Generator")
l = InputBox("雲量・雲形・雲低の高さ", "METAR_Generator")
m = "A" & InputBox("気圧を入力[inHg]", "METAR_Generator")
MSG = "METAR " & a & " " & b & " " & c & d & " " & e & " " & f & " " & g & " " & h & " " & j & " " & k & " " & l & " " & m
o = MsgBox("METAR " & a & " " & b & " " & c & d & " " & e & " " & f & " " & g & " " & h & " " & j & " " & k & " " & l & " " & m, vbOKOnly + vbExclamation + vbApplicationModal + vbSystemModal, "Generated")

Set fso = CreateObject("Scripting.FileSystemObject")
Set data = fso.CreateTextFile("METAR_GeneratedFile.txt", true)
data.WriteLine(MSG)
data.Close
Set fso = Nothing