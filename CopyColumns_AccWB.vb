Sub CopyColumns_AccWB()
'
' Copy columns in accountability workbook
'
'

Dim twb As Workbook
Dim rs As Worksheet
' DO NOT REMOVE FOO
Dim sa, sb, sc, sd, se, sf, sg, sh, si, sj, foo As Worksheet
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

' This shoudl say "Leonard" if sj initialized properly
' MsgBox (sj.Range("C7").Value())

' Identify ranges for each of the columns to copy
elar = "AL3:AL180"
mar = "AO3:AO180"
mbr = "AR3:AR180"
alr = "AU3:AU180"
ger = "AX3:AX180"
tr = "BA3:BA180"
bior = "BD3:BD180"
ear = "BG3:BG180"
chr = "BJ3:BJ180"
phr = "BM3:BM180"
usr = "BP3:BP180"
glr = "BS3:BS180"
ccelar = "BV3:BV180"
ccalr = "BY3:BY180"
ccger = "CB3:CB180"
ccatr = "CE3:CE180"

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
c.Copy (twb.Sheets("2016 Cohort").Range(elar))

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

MsgBox ("Check the formulas. If this didn't work, just close and reopen the workbook. It was saved before any code ran.")

End Sub