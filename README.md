# RTITxtCleaner
Right To Information Request Special Character cleaner
Create a microsoft Excel Sheet

Create button between colums A and F
Enlarge A1 and F1 to enclose the text
button. Click event add the following code
Sub RTIClean()
ActiveSheet.Cells(1, 6).Value = ""
ActiveSheet.Cells(1, 6).Value = RemoveUnAllowedChars(ActiveSheet.Cells(1, 1).Value)
End Sub

Function RemoveUnAllowedChars(Str As String) As String
'updatebyExtendoffice 20160303
    Dim xChars As String
    Dim I, J As Long
    Dim s, c As String
    
    xChars = ",.-_()/@:&?\%abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
   'xChars = ",.-_()/@:&?\%abcdefghijklmnopqrstuvwxyz 0123456789"
    For J = 1 To Len(Str)
        s = Mid$(Str, J, 1)
        If InStr(1, xChars, s, vbBinaryCompare) = 0 Then
       
                         Str = Replace$(Str, s, "")
                          Else
        End If
      
    Next J
    RemoveUnAllowedChars = Str
End Function

Now you can paste the RTI text in A1 and click Button then F1 will contain the cleaned text
