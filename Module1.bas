Attribute VB_Name = "Module1"
'Dichiaro la funzione addetta a mostrare e nascondere il mouse del cursore
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public TV8 As New TrueVision8
Public Scena As New Scene8
Public Schermo As New Screen8
Public TextureFAC As New TextureFactory8
Public Comandi As New InputEngine8
Public Effetti As New GraphicEffect8
Public CD() As Mesh8
Public NCD As Integer
Public NCanzoni As Integer
Public NArtisti As Integer

Type tparamcd
    Texture As String * 100
    Album As String * 50
    Artista As String * 50
    ID As Integer
    Lasciato As Boolean
    
    Canzone(1 To 30) As String * 50
    IDCanzone(1 To 30) As String * 5

    
    X As Single
    Y As Single
    Z As Single
    RZ As Single
    RY As Single
    
End Type
Public Paramcd() As tparamcd

Type TPos
    X As Single
    Y As Single
    Z As Single
    RX As Single
    RY As Single
    RZ As Single
End Type

Public Pos() As TPos

'Dichiaro la struttura contenente i vari tipi di cd
Enum TAccessi
    Amministratore = 1
    utente = 2
End Enum

Type TUtente
    'Dichiaro una variabile che servirà al programma per capire di quale
    'tipo di accesso si dispone
    Accesso As TAccessi
    'Dichiaro una variabile che servirà al programma per salvare il nome
    'dell'utente
    NomeUtente As String
    'Variabile per salvataggio Password utente
    Password As String
End Type

Public InfoUtente As TUtente

'Dichiaro un'array contenente l'elenco dei cd preferiti
Public CDPreferiti(1 To 100) As Integer
'Dichiaro un indice per l'array sopra definito
Public IPref As Integer
'Dichiaro l'oggetto Recordset
Public DB As Database
'Dichiaro l'oggetto Database
Public RS As Recordset
'Dichiaro una variabile che servirà al programma per capire il codice ID
'delle canzoni
Public ICanzone As Integer
Public Aperto As Boolean
