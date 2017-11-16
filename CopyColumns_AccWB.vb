Sub CopyColumns_AccWB()
'
' Copy columns in accountability workbook
'
'

Dim twb As Workbook
Dim rs As Worksheet
' DO NOT REMOVE FOO
Dim sa, sb, sc, sd, se, sf, sg, sh, si, sj, sk, foo As Worksheet
Dim elar, mar, mbr, alr, ger, tr, bior, ear, chr, phr, usr, glr, ccelar, ccalr, ccger, ccatr As String
Dim c As Range

' Only copy from the Accountability workbook
Set twb = Application.Workbooks("Green Tech Cohort Data Collection Workbook.xlsx")

' Save the workbook, in case this doesn't work
twb.Save

' Use the 2006 cohort as the reference sheet, others as secondary sheets
Set rs = twb.Sheets("2006 Cohort")
Set sa = twb.Sheets("2007 Cohort")
Set sb = twb.Sheets("2008 Cohort")
Set sc = twb.Sheets("2009 Cohort")
Set sd = twb.Sheets("2010 Cohort")
Set se = twb.Sheets("2011 Cohort")
Set sf = twb.Sheets("2012 Cohort")
Set sg = twb.Sheets("2013 Cohort")
Set sh = twb.Sheets("2014 Cohort")
Set si = twb.Sheets("2015 Cohort")
Set sj = twb.Sheets("2016 Cohort")
Set sk = twb.Sheets("2017 Cohort")

' This shoudl say "Leonard" if sj initialized properly
' MsgBox (sj.Range("C7").Value())

' Identify ranges for each of the columns to copy
elar = "AM3:AM180"
mar = "AP3:AP180"
mbr = "AS3:AS180"
alr = "AV3:AV180"
ger = "AY3:AY180"
tr = "BB3:BB180"
bior = "BE3:BE180"
ear = "BH3:BH180"
chr = "BK3:BK180"
phr = "BN3:BN180"
usr = "BQ3:BQ180"
glr = "BT3:BT180"
ccelar = "BW3:BW180"
ccalr = "BZ3:BZ180"
ccger = "CC3:CC180"
ccatr = "CF3:CF180"

Set c = rs.Range(elar)
c.Copy (sa.Range(elar))
c.Copy (sb.Range(elar))
c.Copy (sc.Range(elar))
c.Copy (sd.Range(elar))
c.Copy (se.Range(elar))
c.Copy (sf.Range(elar))
c.Copy (sg.Range(elar))
c.Copy (sh.Range(elar))
c.Copy (si.Range(elar))
c.Copy (sj.Range(elar))
c.Copy (sk.Range(elar))
Application.CutCopyMode = False

Set c = rs.Range(mar)
c.Copy (sa.Range(mar))
c.Copy (sb.Range(mar))
c.Copy (sc.Range(mar))
c.Copy (sd.Range(mar))
c.Copy (se.Range(mar))
c.Copy (sf.Range(mar))
c.Copy (sg.Range(mar))
c.Copy (sh.Range(mar))
c.Copy (si.Range(mar))
c.Copy (sj.Range(mar))
c.Copy (sk.Range(mar))
Application.CutCopyMode = False

Set c = rs.Range(mbr)
c.Copy (sa.Range(mbr))
c.Copy (sb.Range(mbr))
c.Copy (sc.Range(mbr))
c.Copy (sd.Range(mbr))
c.Copy (se.Range(mbr))
c.Copy (sf.Range(mbr))
c.Copy (sg.Range(mbr))
c.Copy (sh.Range(mbr))
c.Copy (si.Range(mbr))
c.Copy (sj.Range(mbr))
c.Copy (sk.Range(mbr))
Application.CutCopyMode = False

Set c = rs.Range(alr)
c.Copy (sa.Range(alr))
c.Copy (sb.Range(alr))
c.Copy (sc.Range(alr))
c.Copy (sd.Range(alr))
c.Copy (se.Range(alr))
c.Copy (sf.Range(alr))
c.Copy (sg.Range(alr))
c.Copy (sh.Range(alr))
c.Copy (si.Range(alr))
c.Copy (sj.Range(alr))
c.Copy (sk.Range(alr))
Application.CutCopyMode = False

Set c = rs.Range(ger)
c.Copy (sa.Range(ger))
c.Copy (sb.Range(ger))
c.Copy (sc.Range(ger))
c.Copy (sd.Range(ger))
c.Copy (se.Range(ger))
c.Copy (sf.Range(ger))
c.Copy (sg.Range(ger))
c.Copy (sh.Range(ger))
c.Copy (si.Range(ger))
c.Copy (sj.Range(ger))
c.Copy (sk.Range(ger))
Application.CutCopyMode = False

Set c = rs.Range(tr)
c.Copy (sa.Range(tr))
c.Copy (sb.Range(tr))
c.Copy (sc.Range(tr))
c.Copy (sd.Range(tr))
c.Copy (se.Range(tr))
c.Copy (sf.Range(tr))
c.Copy (sg.Range(tr))
c.Copy (sh.Range(tr))
c.Copy (si.Range(tr))
c.Copy (sj.Range(tr))
c.Copy (sk.Range(tr))
Application.CutCopyMode = False

Set c = rs.Range(bior)
c.Copy (sa.Range(bior))
c.Copy (sb.Range(bior))
c.Copy (sc.Range(bior))
c.Copy (sd.Range(bior))
c.Copy (se.Range(bior))
c.Copy (sf.Range(bior))
c.Copy (sg.Range(bior))
c.Copy (sh.Range(bior))
c.Copy (si.Range(bior))
c.Copy (sj.Range(bior))
c.Copy (sk.Range(bior))
Application.CutCopyMode = False

Set c = rs.Range(ear)
c.Copy (sa.Range(ear))
c.Copy (sb.Range(ear))
c.Copy (sc.Range(ear))
c.Copy (sd.Range(ear))
c.Copy (se.Range(ear))
c.Copy (sf.Range(ear))
c.Copy (sg.Range(ear))
c.Copy (sh.Range(ear))
c.Copy (si.Range(ear))
c.Copy (sj.Range(ear))
c.Copy (sk.Range(ear))
Application.CutCopyMode = False

Set c = rs.Range(chr)
c.Copy (sa.Range(chr))
c.Copy (sb.Range(chr))
c.Copy (sc.Range(chr))
c.Copy (sd.Range(chr))
c.Copy (se.Range(chr))
c.Copy (sf.Range(chr))
c.Copy (sg.Range(chr))
c.Copy (sh.Range(chr))
c.Copy (si.Range(chr))
c.Copy (sj.Range(chr))
c.Copy (sk.Range(chr))
Application.CutCopyMode = False

Set c = rs.Range(phr)
c.Copy (sa.Range(phr))
c.Copy (sb.Range(phr))
c.Copy (sc.Range(phr))
c.Copy (sd.Range(phr))
c.Copy (se.Range(phr))
c.Copy (sf.Range(phr))
c.Copy (sg.Range(phr))
c.Copy (sh.Range(phr))
c.Copy (si.Range(phr))
c.Copy (sj.Range(phr))
c.Copy (sk.Range(phr))
Application.CutCopyMode = False

Set c = rs.Range(usr)
c.Copy (sa.Range(usr))
c.Copy (sb.Range(usr))
c.Copy (sc.Range(usr))
c.Copy (sd.Range(usr))
c.Copy (se.Range(usr))
c.Copy (sf.Range(usr))
c.Copy (sg.Range(usr))
c.Copy (sh.Range(usr))
c.Copy (si.Range(usr))
c.Copy (sj.Range(usr))
c.Copy (sk.Range(usr))
Application.CutCopyMode = False

Set c = rs.Range(glr)
c.Copy (sa.Range(glr))
c.Copy (sb.Range(glr))
c.Copy (sc.Range(glr))
c.Copy (sd.Range(glr))
c.Copy (se.Range(glr))
c.Copy (sf.Range(glr))
c.Copy (sg.Range(glr))
c.Copy (sh.Range(glr))
c.Copy (si.Range(glr))
c.Copy (sj.Range(glr))
c.Copy (sk.Range(glr))
Application.CutCopyMode = False

Set c = rs.Range(ccelar)
c.Copy (sa.Range(ccelar))
c.Copy (sb.Range(ccelar))
c.Copy (sc.Range(ccelar))
c.Copy (sd.Range(ccelar))
c.Copy (se.Range(ccelar))
c.Copy (sf.Range(ccelar))
c.Copy (sg.Range(ccelar))
c.Copy (sh.Range(ccelar))
c.Copy (si.Range(ccelar))
c.Copy (sj.Range(ccelar))
c.Copy (sk.Range(ccelar))
Application.CutCopyMode = False

Set c = rs.Range(ccalr)
c.Copy (sa.Range(ccalr))
c.Copy (sb.Range(ccalr))
c.Copy (sc.Range(ccalr))
c.Copy (sd.Range(ccalr))
c.Copy (se.Range(ccalr))
c.Copy (sf.Range(ccalr))
c.Copy (sg.Range(ccalr))
c.Copy (sh.Range(ccalr))
c.Copy (si.Range(ccalr))
c.Copy (sj.Range(ccalr))
c.Copy (sk.Range(ccalr))
Application.CutCopyMode = False

Set c = rs.Range(ccger)
c.Copy (sa.Range(ccger))
c.Copy (sb.Range(ccger))
c.Copy (sc.Range(ccger))
c.Copy (sd.Range(ccger))
c.Copy (se.Range(ccger))
c.Copy (sf.Range(ccger))
c.Copy (sg.Range(ccger))
c.Copy (sh.Range(ccger))
c.Copy (si.Range(ccger))
c.Copy (sj.Range(ccger))
c.Copy (sk.Range(ccger))
Application.CutCopyMode = False

Set c = rs.Range(ccatr)
c.Copy (sa.Range(ccatr))
c.Copy (sb.Range(ccatr))
c.Copy (sc.Range(ccatr))
c.Copy (sd.Range(ccatr))
c.Copy (se.Range(ccatr))
c.Copy (sf.Range(ccatr))
c.Copy (sg.Range(ccatr))
c.Copy (sh.Range(ccatr))
c.Copy (si.Range(ccatr))
c.Copy (sj.Range(ccatr))
c.Copy (sk.Range(ccatr))
Application.CutCopyMode = False

MsgBox ("Check the formulas. If this didn't work, just close and reopen the workbook. It was saved before any code ran.")

End Sub
