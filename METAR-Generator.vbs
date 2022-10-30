Option Explicit
'Const CREATE_FILE

Dim MSG, a, b, c, d, di, de, ei, e, f, g, h1, h2, h, i, j, k, l, m, n, o
Dim fso, data

a = InputBox("’nˆæ—ª†‚ğ“ü—Í", "METAR_Generator")
b = InputBox("ŠÏ‘ª“ú‚ğ“ü—Í(UTC)[“ú+ŠÔ]", "METAR_Generator") & "Z"
c = InputBox("•—Œü‚ğ“ü—Í(“x)", "METAR_Generator")
di = InputBox("•—‘¬‚ğ“ü—Í(m/s)", "METAR_Generator")
de = di * 1.943844
d = Round(de) & "KT"
e = InputBox("‹’ö‚ğ“ü—Í(m)", "METAR_Generator")
'if ei >= 10000 Then
'	e = 9999
'if ei >= 5000 and ei < 10000 Then
'	e = Round(ei, -3)
'if ei < 5000 Then
'	e = Round(ei, -2)
f = InputBox("Œ»İ‚Ì“V‹C‚ğ“ü—Í", "METAR_Generator")
g = InputBox("‰_—ÊE‰_’ò‚Ì‚‚³‚ğ“ü—Í", "METAR_Generator")
h1 = InputBox("‹C‰·‚ğ“ü—Í", "METAR_Generator")
h2 = InputBox("˜I“_‚ğ“ü—Í", "METAR_Generator")
h = h1 & "/" & h2
j = "Q" & InputBox("‹Cˆ³‚ğ“ü—Í[hPa]", "METAR_Generator")
k = InputBox("‘“àw¦•„?[–³‚¯‚ê‚Î‚±‚Ìæ‚©‚ç‚Í‹ó”’‚ÅOK]", "METAR_Generator")
l = InputBox("‰_—ÊE‰_Œ`E‰_’á‚Ì‚‚³", "METAR_Generator")
m = "A" & InputBox("‹Cˆ³‚ğ“ü—Í[inHg]", "METAR_Generator")
MSG = "METAR " & a & " " & b & " " & c & d & " " & e & " " & f & " " & g & " " & h & " " & j & " " & k & " " & l & " " & m
o = MsgBox("METAR " & a & " " & b & " " & c & d & " " & e & " " & f & " " & g & " " & h & " " & j & " " & k & " " & l & " " & m, vbOKOnly + vbExclamation + vbApplicationModal + vbSystemModal, "Generated")

Set fso = CreateObject("Scripting.FileSystemObject")
Set data = fso.CreateTextFile("METAR_GeneratedFile.txt", true)
data.WriteLine(MSG)
data.Close
Set fso = Nothing