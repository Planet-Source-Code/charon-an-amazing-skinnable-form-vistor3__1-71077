VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Minuteur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************
'**                                                     **
'**  CLASSE REMPLASSANT LE CONTROLE "TIMER"             **
'**                                                     **
'**                                                     **
'**  AUTEUR : Thomas John                               **
'**                                                     **
'*********************************************************

'VARIABLES
Private ActifSz As Boolean
Private IntervalleSz As Long
Private TmrLong As Long

'EVENEMENTS
Event Action()

'DECLENCHEMENT DU MINUTEUR
Public Property Let Actif(Valeur As Boolean)
    
    If Valeur = True Then
        
        If ActifSz <> True Then
            
            'il faut que l'intervalle soit plus grand que zéro
            If IntervalleSz > 0 Then
            
                'on lance le minuteur
                TmrLong = SetTimer(0, 0, IntervalleSz, AddressOf MinuteurProc)
                'on ajoute la référence de notre classe au tableau prévu à cet effet
                AjoutColl ObjPtr(Me), TmrLong
                
            End If
            
        End If
        
    Else
        
        If ActifSz <> False Then
        
            'on stoppe le minuteur
            If TmrLong > 0 Then
                
                'on "tue" le minuteur
                Call KillTimer(0, TmrLong)
                
                'on enlève la référence à notre classe du tableau
                EnleveColl ObjPtr(Me)
                
                TmrLong = 0
                
            End If
            
        End If
        
    End If
    
    ActifSz = Valeur

End Property

Public Property Get Actif() As Boolean

    Actif = ActifSz

End Property

'INTERVALLE
Public Property Let Intervalle(Valeur As Long)

    IntervalleSz = Valeur

End Property

Public Property Get Intervalle() As Long

    Intervalle = IntervalleSz

End Property

'INITIALISATION DE LA CLASSE
Private Sub Class_Initialize()
    
    ActifSz = False

End Sub

'DECHARGEMENT DE LA CLASSE
Private Sub Class_Terminate()
    
    'on stoppe le minuteur
    If TmrLong > 0 Then
        Call KillTimer(0, TmrLong)
        EnleveColl ObjPtr(Me)
    End If
    
End Sub

'ACTION
Friend Sub LancementAction()

    If ActifSz = False Then
        
        Exit Sub
        
    End If
    
    RaiseEvent Action
    
End Sub
