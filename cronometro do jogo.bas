Attribute VB_Name = "Módulo1"
Public Start As Boolean
Sub Cronometro()

'On Error GoTo Erro

Dim T1 As Date, T2 As Date

If Start = True Then

   T1 = Planilha1.Range("F7").Value
   T2 = T1 + VBA.TimeValue("00:00:01") + 1
   
   Planilha1.Range("F7").Value = VBA.Format(T2, "hh:mm:ss")
   
   If Planilha1.Range("F7").Text >= "00:00:12" Then
      MsgBox "Término do período!", vbInformation, "TEMPO"
      Planilha1.Range("F7").Value = TimeValue("00:00:00")
      Star = False
      Exit Sub
      
   End If
   
   Application.OnTime VBA.Now + VBA.TimeValue("00:00:01"), "Cronometro"
   
    
End If


Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "CRONÔMETRO"


End Sub

'A função RND vai retornar o valor entre 0 e 1 da cotação

Sub Teste_01()
    Range("C6").Value = Rnd

End Sub
