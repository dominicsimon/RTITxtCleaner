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
