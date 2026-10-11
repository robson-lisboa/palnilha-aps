VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    APS_AoAbrir
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    APS_PararTick
End Sub

' Alteracoes feitas a mao: 02_Operacoes passa pela cascata; 1_Planejamento!A1 (Mes do Planejamento)
' e normalizado para o dia 1 do mes.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.Name = APS_ABA_OPS Then APS_OpsAlterada Target
    If Sh.Name = APS_ABA_PLAN Then APS_MesAlterado Target
End Sub

' Ao voltar para o planejamento, confere se algo mudou (ex.: linha excluida) e redesenha.
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If Sh.Name = APS_ABA_PLAN Then APS_SincronizarSeMudou
End Sub

