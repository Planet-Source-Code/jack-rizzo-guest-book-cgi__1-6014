Attribute VB_Name = "Main2"
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
Dim FirstRec As Long
Dim LastRec As Long
Dim OutMsg As String
Dim TmpMsg As String
Dim RecNr As Long
Dim i As Long
Dim n As Long
Dim xbook As BookInfo
DBSName = GetCgiValue("dbs")
WallPaper = GetCgiValue("wp")
SendHeader "Guest Book Entries"
wpstring = "<BODY BACKGROUND=""" & WallPaper & """" & " TEXT=""#00ffff"" LINK=""#00ffff"" VLINK=""#ffff00"">"
Send wpstring
DBSName = DBSName & ".dbs"
RecNr = Val(GetCgiValue("recnr")) + 1
RecCount = 0
irec = 1
ix = FreeFile
Open DBSName For Random Access Read Write Lock Write As #ix Len = Len(Book)
Get #ix, irec, xbook
RecCount = Val(xbook.Iname)
If RecCount < 6 Then
   LastRec = RecCount + 1
Else
   If (RecNr + 5) > RecCount Then
      LastRec = RecCount + 1
   Else
      LastRec = RecNr + 4
   End If
End If
FirstRec = RecNr
Send "<BR><BR><CENTER><FONT SIZE=""+2"">Guest Book Entries</FONT><BR>"
Send "There are " & Str(RecCount) & " Entries in the Guest Book</CENTER><BR><BR>"
Send "<TABLE WIDTH=""720"" BORDER=""0"" CELLSPACING=""2"" CELLPADDING=""0"">"
Send "<TD WIDTH=""18%"">"
Send "</TD>"
Send "<TD WIDTH=""82%"">"
For i = FirstRec To LastRec
   Get #ix, i, xbook
   Send "<P>Entry Nr: " & Str(i - 1) & "<BR>"
   If Left(xbook.Iname, 1) = "*" Then
      Send "Deleted!" < BR > ""
      GoTo bypass
   End If
   Send "Name: " & Trim(xbook.Iname) & "<BR>"
   Send "Email: " & Trim(xbook.IEmail) & "<BR>"
   Send "Resident State: " & UCase(Trim(xbook.IState)) & "<BR>"
   Send "Date of Entry: " & Trim(xbook.Idate) & "<BR>"
   TmpMsg = Trim(xbook.Icomments)
   n = 1
   OutMsg = ""
   For k = n To Len(TmpMsg)
      If Mid(TmpMsg, k, 2) = vbCrLf Then
         OutMsg = OutMsg & "<BR>"
         k = k + 1
      Else
         OutMsg = OutMsg & Mid(TmpMsg, k, 1)
      End If
   Next k
   Send "Remarks: " & OutMsg & "<BR><BR>"
bypass:
Next i
Send "</TD>"
Send "</TR>"
Send "</TABLE><BR>"
Send "<CENTER><P>"
TmpMsg = "recnr=" & FirstRec - 6
If i > 7 Then
    Send "<A HREF=""http://sbnsoftware.com/cgi_bin/guest2.exe?dbs=becca&wp=http://beccaanderson.com/flag4.jpg&def=http://beccaanderson.com/default.htm&" & TmpMsg & """ METHOD=""POST"" ENCTYPE=""x-www-form-urlencoded"">[Previous 5]</A>&nbsp;&nbsp;"
End If
TmpMsg = "recnr=" & FirstRec + 4
If FirstRec + 2 < LastRec Then
    Send "<A HREF=""http://sbnsoftware.com/cgi_bin/guest2.exe?dbs=becca&wp=http://beccaanderson.com/flag4.jpg&def=http://beccaanderson.com/default.htm&" & TmpMsg & """ METHOD=""POST"" ENCTYPE=""x-www-form-urlencoded"">[Next 5]</A>"
End If
Send "<BR><A HREF=""" & DEFX & """>[Home]<BR><BR>"
Send "</CENTER><BR><BR>"
SendFooter
      
End
End Sub
