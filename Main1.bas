Attribute VB_Name = "Main1"
Type BookInfo
   Iname As String * 40
   IEOL1 As String * 2
   IEmail As String * 40
   IEOL2 As String * 2
   IState As String * 2
   IEOL3 As String * 2
   Icomments As String * 2048
   IEOL4 As String * 2
   Idate As String * 20
   IEOL5 As String * 2
End Type
Public Book As BookInfo
Public Sub Cgi_Main()
Dim ix As Long
Dim irec As Long
Dim RecCount As Long
Dim DBSName As String
Dim WallPaper As String
Dim wpstring As String
Dim xbook As BookInfo
DBSName = GetCgiValue("dbs")
WallPaper = GetCgiValue("wp")
wpstring = "<BODY BACKGROUND=""" & WallPaper & """" & " TEXT=""#00ffff"" LINK=""#00ffff"" VLINK=""#ffff00"">"
Book.Iname = GetCgiValue("Name")
Book.IEmail = GetCgiValue("Email")
Book.IState = GetCgiValue("State")
Book.Icomments = GetCgiValue("Comments")
Book.Idate = Now
If Trim(Book.Iname) = "" Or Trim(Book.Icomments) = "" Or Trim(Book.IState) = "" Then
   SendHeader "Guest Book Entry ERROR"
   Send wpstring
   Send "<BR><BR><CENTER><FONT SIZE=""+2"">You must enter a Name, Resident State and Comments!</FONT><BR><BR>"
   Send "<A HREF=""" & DEFX & """>Home</A></B></CENTER>"
   SendFooter
   Exit Sub
End If
If DBSName <> "" Then
   SendHeader "Guest Book Entry Processed"
   Send " "
   Send wpstring
   Send " "
   Send "<BR><BR><CENTER><FONT SIZE=""+2"">Thank you for Signing the Guest Book</FONT><BR><BR>"
   Send "<A HREF="""
   Send DEFX & """"
   Send ">Home</A></B></CENTER>"
   SendFooter
Else
   SendHeader "Gustbook Entry ERROR"
   Send wpstring
   Send "<BR><BR><Center>Unable to find Guest book File!<BR><BR>"
   Send "<A HREF=""" & DEFX & """>Home</A></B></CENTER>"
   SendFooter
   Exit Sub
End If
Book.IEOL1 = vbCrLf
Book.IEOL2 = vbCrLf
Book.IEOL3 = vbCrLf
Book.IEOL4 = vbCrLf
Book.IEOL5 = vbCrLf
DBSName = DBSName & ".dbs"
irec = 1
RecCount = 0
ix = FreeFile
Open DBSName For Random Access Read Write Lock Write As #ix Len = Len(Book)
Get #ix, irec, xbook
If Asc(xbook.Iname) = 0 Then
    xbook.Iname = "1"
Else
    RecCount = Val(Trim(xbook.Iname)) + 1
    xbook.Iname = Str(RecCount)
End If
Put #ix, irec, xbook
irec = Val(xbook.Iname) + 1
Put #ix, irec, Book
Close #ix
End
End Sub
