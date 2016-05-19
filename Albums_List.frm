VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "Jukebox"
   ClientHeight    =   11520
   ClientLeft      =   180
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Check_Disc 
      Interval        =   100
      Left            =   480
      Top             =   11040
   End
   Begin MSWinsockLib.Winsock Mittente 
      Left            =   4200
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Dedica 
      Left            =   3720
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   313
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS Computer 
      Height          =   135
      Left            =   1920
      OleObjectBlob   =   "Albums_List.frx":0000
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   11040
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   2160
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MediaPlayerCtl.MediaPlayer WM 
      Height          =   30
      Left            =   15360
      TabIndex        =   0
      Top             =   11520
      Visible         =   0   'False
      Width           =   135
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   -1  'True
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -630
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Motore_Sonoro As SoundEngine8
Dim Suoni As Sounds8
'Dichiaro un'array che servirà a contenere i nomi delli traccie MP3 che si andranno ad analizzare
'private del loro numero
Dim Traccia(1 To 60) As String
'Definisco un indice per l'array appena definito
Dim NumeroTraccia As Integer
Dim TotFiles As Integer
'Dichiaro un'array che servirà a contenere il testo della canzone di massimo 300 righe
Dim Lyrics(1 To 300) As String
'Dichiaro una variabile che servirà al programma per capire se è stato cancellato
'il testo della canzone riprodotta precedentemente
Dim Cancellato As Boolean
'Dichiaro una variabile che servirà per capire di quante righe è composta il testo della
'canzone in esecuzione
Dim NRigheL As Integer
'Dichiaro due variabili che serviranno a creare l'effetto blink delle due freccie
Dim Blink As Single
Dim Direzione As String
Dim Preso As Boolean
Dim FirstCicle As Boolean
Dim CDCorrente As Integer
Dim Rotazione As Single

Dim CDMenù As Boolean
Dim MenùImporta As Boolean
Dim ModificaPreferiti As Boolean

'Dichiaro una variabile che servirà a contenere il numero temporaneo della canzone richiesta
Dim CanzoneRichiesta As String
'Dichiaro una variabile che servirà a contenere il nome della canzone e del gruppo richiesto
Dim InfoCanzone As String
'Dichiaro una variabile che servirà a contenere il nome del file contenente il testo della
'canzone in esecuzione
Dim FileTestoCanzone As String
'Dichiaro una variabile che servirà al programma per capire quanto tempo è passato dall'ultima
'pressione di un tasto sul PAD numerico
Dim TimePAD As Single
'Dichiaro un array che servirà a contenere la coda delle canzone richieste
Dim Buffer(1 To 100) As String
'Dichiaro l'indice che verrà utilizzato dall'array soprastante
Dim IB As Integer
'Dichiaro un'altra variabile che servirà al programma per capire quale canzone dovrà
'eseguire
Dim IC As Integer
Dim Z As Integer
Dim Esecuzione_Programma As Boolean
'Dichiaro due variabili che serviranno per impostare l'avanzamento delle due barre
'progressive per l'importazione di nuovi cd
Dim AvanzTOTALE As Integer
Dim AvanzPARZIALE As Integer
Dim TmpTexture As String

Dim TextureFront As Boolean
Dim TextureBack As Boolean

Dim TmpArtista As String
Dim TmpTitolo As String
Dim TrackNumber As String

Dim LastKeyPressed As CONST_TV8_KEY
Dim Released As Boolean
Dim OKFLASH As Boolean
Dim OKMENUCD As Boolean

Private Sub Check_Disc_Timer()
    'Viene verificato se è stato inserito un CD e,in caso affermativo, viene controllato
    'se questo è vergine...
    On Error Resume Next
    Open Log_IN.SelectedDrive.DriveLetter + ":\testcd" For Output As #1
    Close #1
    'Se non si è verificato nessun errore, vuol dire che il cd è stato inserito,
    'quindi...
    If Err.Number = 71 Then
        Aperto = True
    'Altrimenti, il cd è vuoto, quindi...
    Else
        If Aperto = True Then
            'Aperto = False
            Distruggi
            Log_IN.Show
            'Unload Me
        End If
    End If

End Sub

Private Sub Dedica_ConnectionRequest(ByVal requestID As Long)
    Dim RH As String
    'Viene accettata la connessione remota
    Mittente.Accept requestID
    Computer.CurrentMode = 4
    Computer.Speak "A new user has connected!"
    
End Sub

Private Sub Form_Load()
    'Chiamo la funzione addetta all'inizializzazione del main menù del software
    'interamente in 3D
    Inizializza_3D
    'Chiamata alla funzione di inizializzazione variabili d'ambiente
    Inizializza_Variabili
    'Viene richiamata la funzione di creazione CD
    CreaCD
    'Chiamo la funzione che permetterà di creare il menù principale
    Main_Menù
    'Chiamata alla funzione di inizializzazione connessione
    Inizializza_Connessione
    'Viene Richiamto il Loop principale del software "JukeBox" il quale permetterà di
    'realizzare per l'appunto tutti i menù che compongono il programma sottoscritto
    Ciclo_3D
End Sub

Private Sub Inizializza_Variabili()
    'Inizializzo le variabile che permetteranno di creare un effetto blink sulle
    'due freccie posizionate ai lati del CD centrale
    Blink = 1
    Direzione = "Down"
    Lasciato = True
    FirstCicle = True
    CDCorrente = 1
    IB = 1
    IC = 0
    SelectedOperator = 1
    Released = True
    Esecuzione_Programma = True
    Computer.CurrentMode = 4
    Set Suoni = New Sounds8
    Set Motore_Sonoro = New SoundEngine8
    'Carico gli effetti sonori del software
    Suoni.AddFile App.Path & "\Sounds\Keyb.wav", "PAD"
    Suoni.AddFile App.Path & "\Sounds\Error.wav", "ERROR"
    Suoni.AddFile App.Path & "\Sounds\OK.wav", "OK"
End Sub

Private Sub Inizializza_3D()
    'Inizializzo il motore 3D a schermo interno (1024x768)
    'TV8.Init3DFullscreen 1024, 768, 32
    TV8.Init3DWindowedMode Me.hwnd
    Me.Show
    'Il MipMapping è una particolare funzione 3D che mi permetterà di migliorare la grafica nella modalità
    'Anteprima3D
    Scena.EnableMipMapping True
    Esecuzione_Programma = True
    Scena.SetShadeMode TV_SHADEMODE_PHONG
    TextureFAC.SetTextureMode TV_TEXTUREMODE_32BITS
    Scena.SetTextureFilter TV_FILTER_ANISOTROPIC
    TV8.EnableAntialising True
End Sub

Private Sub Ciclo_3D()
    Do
        DoEvents
        'Ripulisce il contenuto dell'oggetto TV8
        TV8.Clear
        'Richiamo la funzione addetta al controllo del buffer del jukebox
        Controlla_Buffer
        If CanzoneRichiesta <> "" Then
            Timer1.Interval = 1
            If Comandi.IsKeyPressed(LastKeyPressed) = False Then
                Released = True
            End If
        Else
            Timer1.Interval = 100
            If Comandi.IsKeyPressed(LastKeyPressed) = False Then
                Released = True
            End If
        End If
        If CDMenù = False And MenùImporta = False Then
        
            'Richiamo la funzione che,renderizzando tutte le mesh presenti nella scena
            'permetterà la realizzazione del main menù contenente tutti i CD in versione 3D
            Scena.RenderAllMeshes
            'Richiamo la funzione addetta alla visualizzazione di tutti gli elementi 2D presenti
            'all'interno del menù
            Attiva_Elementi_2D
        ElseIf CDMenù = True And MenùImporta = False Then
            Disegna_Grafica
            'Renderizzo il CD corrente
            CD(CDCorrente).Render
            Disegna_MenùCD
        End If
        If MenùImporta = True Then
            Disegna_Grafica
            Disegna_Menù_Importa
        End If
        'Se il CD corrente non è stato ancora portato in primo piano
        If Preso = False Then
            'Se è la prima volta che si entra nella funzione Prendi_CD per il CDCorrente
            If FirstCicle = True Then
                Trasferisci_Attributi CDCorrente
            End If
            'Richiamo la funzione che permetterà di prelevare il CDCorrente e portarlo
            'in primo piano sullo schermo
            Prendi_CD CDCorrente
        Else
            'Incremento la variabile Rotazione di 0.01 in modo che l'annello 3D non stia mai fermo ma simuli
            'continuamente un effetto di rotazione
            Rotazione = Rotazione + 0.01 + TV8.TimeElapsed / 20
            'Ruoto in senso orario il cd posto in primo piano sullo schermo
            CD(CDCorrente).SetRotation 0, 180 + Rotazione, 0
        End If
        For Z = (CDCorrente - 5) To CDCorrente + 5
            On Error Resume Next
            If Paramcd(Z).Lasciato = False Then
                If Err.Number = 0 Then
                    Posa_CD Z
                End If
            End If
        Next
        For Z = 1 To 5
            On Error Resume Next
            If Paramcd(Z).Lasciato = False Then
                If Err.Number = 0 Then
                    Posa_CD Z
                End If
            End If
        Next
        For Z = (NCD - 5) To NCD + 5
            On Error Resume Next
            If Paramcd(Z).Lasciato = False Then
                If Err.Number = 0 Then
                    Posa_CD Z
                End If
            End If
        Next
        'Per risolvere un piccolo BUG controllo il primo CD e l'ultimo caricato
        If Paramcd(1).Lasciato = False Then Posa_CD 1
        If Paramcd(NCD).Lasciato = False Then Posa_CD NCD
        'Renderizza tutto il contenuto dell'oggetto TV8 su schermo
        TV8.RenderToScreen
    'Tutto questo avviene finchè la variabile continua_animazione non assumerà il valore da false
    Loop Until Esecuzione_Programma = False
    
    Set Schermo = Nothing
    Set Scena = Nothing
    For Z = 1 To NCD
        Set CD(Z) = Nothing
    Next
    Set TV8 = Nothing
    End
End Sub

Private Sub Controlla_Comandi()
    Dim I As Integer
    If Released = True Then
    If Comandi.IsKeyPressed(TV_KEY_ESCAPE) = True Then Esecuzione_Programma = False
    
    'Alla pressione del pulsante numero 6 del tastierino numerico...
    If CDMenù = False And Comandi.IsKeyPressed(TV_KEY_NUMPAD3) = True Then
        'Se non si è già raggiunto l'ultimo CD
        If CDCorrente < NCD + 1 Then
            'Passo al cd successivo incrementando la variabile CDCorrente
            CDCorrente = CDCorrente + 1
        ElseIf CDCorrente = NCD + 1 Then
            CDCorrente = 1
        End If
        If CDCorrente - 1 = 0 Then
            Paramcd(NCD + 1).Lasciato = False
        Else
            Paramcd(CDCorrente - 1).Lasciato = False
            'Reinizializzo le due variabile che permetteranno,al ciclo successivo,
            'di richiamare alla funzione PrendiCD
            Preso = False
            FirstCicle = True
        End If
        'Salvo l'ultimo pulsante premuto
        LastKeyPressed = TV_KEY_NUMPAD3
    End If
    'Alla pressione del pulsante numero 4 del tastierino numerico...
    If CDMenù = False And Comandi.IsKeyPressed(TV_KEY_NUMPAD1) = True Then
        'Se non si è già raggiunto il primo CD
        If CDCorrente > 1 Then
            'Passo al cd precedente decrementando la variabile CDCorrente
            CDCorrente = CDCorrente - 1
        ElseIf CDCorrente = 1 Then
            CDCorrente = NCD + 1
        End If
        If (CDCorrente + 1) > NCD + 1 Then
            Paramcd(1).Lasciato = False
        Else
            Paramcd(CDCorrente + 1).Lasciato = False
            'Reinizializzo le due variabile che permetteranno,al ciclo successivo,
            'di richiamare alla funzione PrendiCD
            Preso = False
            FirstCicle = True
        End If
        'Salvo l'ultimo pulsante premuto
        LastKeyPressed = TV_KEY_NUMPAD1
    'Altrimenti se è stato premuto un qualsiasi altro tasto dal PAD numerico...
    ElseIf Comandi.IsKeyPressed(TV_KEY_NUMPAD4) = True Or Comandi.IsKeyPressed(TV_KEY_NUMPAD5) = True Or Comandi.IsKeyPressed(TV_KEY_NUMPAD6) = True Or Comandi.IsKeyPressed(TV_KEY_NUMPAD7) = True Or Comandi.IsKeyPressed(TV_KEY_NUMPAD8) = True Or Comandi.IsKeyPressed(TV_KEY_NUMPAD9) = True Or Comandi.IsKeyPressed(TV_KEY_NUMPADSLASH) = True Or Comandi.IsKeyPressed(TV_KEY_NUMPADSTAR) = True Or Comandi.IsKeyPressed(TV_KEY_NUMPAD2) = True Or Comandi.IsKeyPressed(TV_KEY_NUMLOCK) = True Then
        If Comandi.IsKeyPressed(TV_KEY_NUMPAD4) = True Then
            CanzoneRichiesta = CanzoneRichiesta + "7"
            'Salvo l'ultimo pulsante premuto
            LastKeyPressed = TV_KEY_NUMPAD4
        ElseIf Comandi.IsKeyPressed(TV_KEY_NUMPAD5) = True Then
            CanzoneRichiesta = CanzoneRichiesta + "8"
            'Salvo l'ultimo pulsante premuto
            LastKeyPressed = TV_KEY_NUMPAD5
        ElseIf Comandi.IsKeyPressed(TV_KEY_NUMPAD6) = True Then
            CanzoneRichiesta = CanzoneRichiesta + "9"
            'Salvo l'ultimo pulsante premuto
            LastKeyPressed = TV_KEY_NUMPAD6
        ElseIf Comandi.IsKeyPressed(TV_KEY_NUMPAD7) = True Then
            CanzoneRichiesta = CanzoneRichiesta + "4"
            'Salvo l'ultimo pulsante premuto
            LastKeyPressed = TV_KEY_NUMPAD7
        ElseIf Comandi.IsKeyPressed(TV_KEY_NUMPAD8) = True Then
            CanzoneRichiesta = CanzoneRichiesta + "5"
            'Salvo l'ultimo pulsante premuto
            LastKeyPressed = TV_KEY_NUMPAD8
        ElseIf Comandi.IsKeyPressed(TV_KEY_NUMPAD9) = True Then
            CanzoneRichiesta = CanzoneRichiesta + "6"
            'Salvo l'ultimo pulsante premuto
            LastKeyPressed = TV_KEY_NUMPAD9
        ElseIf Comandi.IsKeyPressed(TV_KEY_NUMPADSLASH) = True Then
            CanzoneRichiesta = CanzoneRichiesta + "2"
            'Salvo l'ultimo pulsante premuto
            LastKeyPressed = TV_KEY_NUMPADSLASH
        ElseIf Comandi.IsKeyPressed(TV_KEY_NUMPADSTAR) = True Then
            CanzoneRichiesta = CanzoneRichiesta + "3"
            'Salvo l'ultimo pulsante premuto
            LastKeyPressed = TV_KEY_NUMPADSTAR
        ElseIf Comandi.IsKeyPressed(TV_KEY_NUMPAD2) = True Then
            CanzoneRichiesta = CanzoneRichiesta + "0"
            'Salvo l'ultimo pulsante premuto
             LastKeyPressed = TV_KEY_NUMPAD2
        ElseIf Comandi.IsKeyPressed(TV_KEY_NUMLOCK) = True Then
            CanzoneRichiesta = CanzoneRichiesta + "1"
            'Salvo l'ultimo pulsante premuto
            LastKeyPressed = TV_KEY_NUMLOCK
        End If
        'Suoni("PAD").Play
        TimePAD = 0
    End If
    
    If Comandi.IsKeyPressed(TV_KEY_NUMPAD0) Then
        If CanzoneRichiesta <> "" And CanzoneRichiesta <> "01" And CanzoneRichiesta <> "02" Then
            Trova_Canzone CanzoneRichiesta
        ElseIf CanzoneRichiesta = "" Then
            If CDCorrente <> NCD + 1 Then
                If CDMenù = True Then
                    CDMenù = False
                    OKMENUCD = False
                Else
                    'Imposto a True la variabile CDMenù, la quale farà capire al programma che
                    'dovrà creare su schermo il menù di informazione CD
                    CDMenù = True
                End If
            'Se il CD scelto è quello che permette di importare uno o più cd, allora...
            ElseIf CDCorrente = NCD + 1 Then
                If MenùImporta = False Then
                    'Se si è loggati come amministratore,allora...
                    If InfoUtente.Accesso = Amministratore Then
                        MenùImporta = True
                        Disegna_Menù_Importa
                        'Imposto la directory iniziale all'interno della quale recuperare i file MP3
                        Dir1.Path = "D:\CD\"
                        'Richiamo la funzione addetta al conteggio dei file complessivi
                        Imposta_Barra_Avanzamento
                        'Chiamo la funzione addetto al Download dei file MP3 che compongono i vari CD da aggiungere
                        'al jubox, da CD a Hard Disk
                        Copia_Files
                    'Altrimenti se si è loggati come utente....
                    Else
                        'Setto la variabile ModificaPreferiti a valore booleano TRUE
                        ModificaPreferiti = True
                        'Chiamata alla funzione di abilitazione di tutti i cd presenti all'interno
                        'del jukebox
                        Attiva_CD
                    End If
                Else
                    MenùImporta = False
                End If
            End If
        End If
        If CanzoneRichiesta = "01" Then
            Randomize
            CanzoneRichiesta = Fix(NCanzoni * Rnd) + 1
            Trova_Canzone CanzoneRichiesta
        ElseIf CanzoneRichiesta = "02" Then
            WM.Stop
        End If
        CanzoneRichiesta = ""
        'Salvo l'ultimo pulsante premuto
        LastKeyPressed = TV_KEY_NUMPAD0
    End If
    Released = False
    End If
End Sub

Private Sub Attiva_Elementi_2D()
    Disegna_Grafica
    Schermo.DrawFilledBox 1024 / 2 - 170, 768 / 2 + 200, 1024 / 2 + 180, 768 / 2 + 260, RGBA(0, 0, 0.5, 1)
    Schermo.DrawBox 1024 / 2 - 171, 768 / 2 + 199, 1024 / 2 + 181, 768 / 2 + 261, RGBA(0, 0, 0, 1)
    Schermo.DrawBox 1024 / 2 - 172, 768 / 2 + 200, 1024 / 2 + 182, 768 / 2 + 262, RGBA(0, 0, 0, 1)
    
    'Schermo.DrawLine 1024 / 2 + 182, 768 / 2 + 230, 1024 / 2 + 230, 768 / 2 + 230, RGBA(0, 0, 0, 1)
    'Schermo.DrawLine 1024 / 2 + 182, 768 / 2 + 231, 1024 / 2 + 230, 768 / 2 + 231, RGBA(0, 0, 0, 1)
    
    'Schermo.DrawLine 1024 / 2 - 172, 768 / 2 + 230, 1024 / 2 - 230, 768 / 2 + 230, RGBA(0, 0, 0, 1)
    'Schermo.DrawLine 1024 / 2 - 172, 768 / 2 + 231, 1024 / 2 - 230, 768 / 2 + 231, RGBA(0, 0, 0, 1)
    
    'Box freccia sinistra
    'Schermo.DrawFilledBox 1024 / 2 - 230, 768 / 2 + 215, 1024 / 2 - 275, 768 / 2 + 245, RGBA(0, 0, 0.5, 1)

    'Se non è selezionato il cd che permette di importare nuovo/i cd, allora...
    If CDCorrente <> NCD + 1 Then
        Schermo.DrawText "Artista:", 1024 / 2 - 160, 768 / 2 + 205, RGBA(1, 1, 0, 1), "Carattere2"
        Schermo.DrawText "Album:", 1024 / 2 - 160, 768 / 2 + 230, RGBA(1, 1, 0, 1), "Carattere2"
    'Se invece è selezionato proprio quel CD
    Else
        'Se è stato effettuato l'accesso come amministratore, allora...
        If InfoUtente.Accesso = Amministratore Then
            Schermo.DrawText "Importa nuovo/i CD", 1024 / 2 - 110, 768 / 2 + 217.5, RGBA(1, 1, 1, 1), "Carattere2"
        'Altrimenti è stato effettuato l'accesso come utente, quindi...
        Else
            Schermo.DrawText "Modifica Jukebox AccessCD", 1024 / 2 - 140, 768 / 2 + 217.5, RGBA(1, 1, 1, 1), "Carattere2"
        End If
    End If
    'Scrivo nell'apposito riquadro, il nome dell'artista e il titolo dell'album
    Schermo.DrawText Paramcd(CDCorrente).Artista, 1024 / 2 - 70, 768 / 2 + 205, RGBA(1, 1, 1, 1), "Carattere4"
    Schermo.DrawText Paramcd(CDCorrente).Album, 1024 / 2 - 70, 768 / 2 + 230, RGBA(1, 1, 1, 1), "Carattere4"
    '----------------------------------------------------------------------------------------
    'Inserisco l'immagine delle 2 freccie ai rispettivi lati del CD centrale
    Schermo.DrawTexture GetTex("Freccia"), 330, 600, 300, 630, RGBA(1, 1, 0, Blink)
    Schermo.DrawTexture GetTex("Freccia"), 700, 600, 730, 630, RGBA(1, 1, 0, Blink)
    'Schermo.DrawText "Precedente", 250, 340, RGBA(1, 1, 1, Blink), "Carattere5"
    'Schermo.DrawText "Successivo", 680, 340, RGBA(1, 1, 1, Blink), "Carattere5"
    
End Sub

Private Sub Disegna_Grafica()
    'Traccio le bande nere stile 16:9
    Schermo.DrawFilledBox 0.01, 0.01, 1023.9, 120, RGBA(0, 0, 0, 1)
    Schermo.DrawFilledBox 0.01, 668, 1023.9, 768.9, RGBA(0, 0, 0, 1)
    'Scrivo la data
    Schermo.DrawText Date, 800, 30, RGBA(1, 1, 1, 1), "Carattere1"
    'Scrivo l'ora
    Schermo.DrawText Time, 855, 70, RGBA(1, 1, 1, 1), "Carattere1"
    '--------------------------------------------------------------------------------------
    'Disegno il box posizionato in basso a destra dello schermo e ne inserisco le scritte
    'di statistiche
    '--------------------------------------------------------------------------------------
    'Schermo.DrawFilledColorBox 650, 670, 1023, 767, RGBA(0.5, 0, 0, 0.7), RGBA(0.7, 0, 0, 1), RGBA(0.8, 0, 0, 1), RGBA(1, 0.2, 0, 1)
    'Se non si è in modalità modifica lista CD preferiti, allora...
    If ModificaPreferiti = False Then
        Schermo.DrawText "Totale CD:", 825, 680, RGBA(1, 1, 1, 1), "Carattere3"
        Schermo.DrawText "Totale Canzoni:", 772, 705, RGBA(1, 1, 1, 1), "Carattere3"
        Schermo.DrawText "Numero Artisti:", 780, 730, RGBA(1, 1, 1, 1), "Carattere3"
    
        Schermo.DrawText CStr(NCD), 950, 680, RGBA(1, 1, 0, 1), "Carattere3"
        Schermo.DrawText CStr(NCanzoni), 950, 705, RGBA(1, 1, 0, 1), "Carattere3"
        Schermo.DrawText CStr(NArtisti), 950, 730, RGBA(1, 1, 0, 1), "Carattere3"
        '--------------------------------------------------------------------------------------
        ''Schermo.DrawFilledColorBox 0.1, 670, 400, 767, RGBA(0.5, 0, 0, 0.7), RGBA(0.7, 0, 0, 1), RGBA(0.8, 0, 0, 1), RGBA(1, 0.2, 0, 1)
        'Schermo.DrawTexture GetTex("Screen"), 0.1, 655, 400, 767
        If WM.PlayState <> 2 Then
            Schermo.DrawText "...Nessuna canzone caricata...", 70, 710, RGBA(1, 1, 1, Blink), "Carattere8"
        End If
        Schermo.DrawText "State ascoltando:", 100, 680, RGBA(1, 1, 1, 1), "Carattere3"
        'Se il JukeBox stà eseguendo una qualche canzone allora...
        If WM.PlayState = 2 Then
            'Scrivo su schermo il titolo e l'artista della canzone che è in esecuzione
            Schermo.DrawText InfoCanzone, 20, 715, RGBA(0, 1, 0, 1), "Carattere7"
        End If
        'Schermo.DrawFilledColorBox 413, 680, 637, 757, RGBA(0, 0, 0.8, 0.7), RGBA(0, 0, 0.3, 1), RGBA(0, 0, 0.5, 1), RGBA(0, 0, 0.7, 1)
        Schermo.DrawText "Canzone richiesta", 433, 680, RGBA(1, 1, 1, 1), "Carattere3"
        Schermo.DrawText CanzoneRichiesta, 500, 715, RGBA(0.6, 1, 0.2, 1), "Carattere3"
    End If
    'Se è passato un po di tempo dall'ultima pressione di un tasto sul tastierino numerico,allora...
    If TimePAD >= 30 Then
        'Azzero il contatore
        TimePAD = 0
        'Azzero la variabile contenente il numero della canzone richiesta
        CanzoneRichiesta = ""
    'Altrimenti...
    Else
        'Continuo a contare il tempo trascorso
        TimePAD = TimePAD + 0.1
    End If
    'Scrivi_Testo_Canzone
End Sub

Private Sub Disegna_MenùCD()
    Dim I As Integer
    Dim X As Long
    Dim Y As Long
    Dim ValTesto As String
    'Scrivo il titolo del cd
    Schermo.DrawText Trim(Paramcd(CDCorrente).Artista) + " - " + Trim(Paramcd(CDCorrente).Album), 300, 130, RGBA(0, 1, 0, 1), "Carattere3"
    Y = 200
    X = 120
    For I = 1 To 30
        If I = 16 Then
            Y = 200
            X = 650
        End If
        If Trim(Paramcd(CDCorrente).Canzone(I)) <> "\" Then
            Schermo.DrawText Trim(Paramcd(CDCorrente).IDCanzone(I)), X - 60, Y, RGBA(0, 1, 0, 1), "Carattere3"
            Schermo.DrawText Trim(Paramcd(CDCorrente).Canzone(I)), X, Y, RGBA(1, 1, 1, 1), "Carattere3"
        End If
        Y = Y + 20
    Next
    End Sub

Private Sub Main_Menù()
    Schermo.CreateUserFont "Carattere1", "ChickenScratch", 30, False, False, False
    Schermo.CreateUserFont "Carattere2", "GlaserSteD", 15, False, False, False
    Schermo.CreateUserFont "Carattere3", "ChickenScratch", 20, False, False, False
    Schermo.CreateUserFont "Carattere4", "Iron Maiden", 15, False, False, False
    Schermo.CreateUserFont "Carattere5", "Verdana", 8, True, False, False
    Schermo.CreateUserFont "Carattere6", "ChickenScratch", 15, False, False, False
    Schermo.CreateUserFont "Carattere7", "Arial Black", 12, False, False, False
    Schermo.CreateUserFont "Carattere8", "Verdana", 10, True, False, False
    Schermo.CreateUserFont "Carattere9", "Verdana", 8, True, False, False
End Sub

Private Sub CreaCD()
    Dim I As Integer
    Dim K As Integer
    Dim TV As Boolean
    Dim Num As Integer
    Dim Num2 As Integer
    Dim Num3 As Integer
    Dim Num4 As Single
    For I = 1 To NCD + 1
        Set CD(I) = Scena.CreateMeshBuilder("Album" + CStr(I))
        CD(I).Load3DsMesh App.Path & "\CDCover.3ds"
        TextureFAC.LoadTexture App.Path & Paramcd(I).Texture, "T" + CStr(I)
        CD(I).SetTexture GetTex("T" + CStr(I))
    Next
    Set CD(NCD + 1) = Scena.CreateMeshBuilder("Album" + CStr(I))
    CD(NCD + 1).Load3DsMesh App.Path & "\CD.3ds"
    'TextureFAC.LoadTexture App.Path & Paramcd(NCD).Texture, "NewCD"
    'CD(NCD).SetTexture GetTex("NewCD")
    For I = 1 To NCD + 1
        Randomize
        Num = Fix(2 * Rnd) + 1
        Num3 = Fix(2 * Rnd) + 1
        If Num = 1 Then
            Num2 = -1
        Else
            Num2 = 1
        End If
        If Num3 = 1 Then
            Paramcd(I).RY = 180
        Else
            Paramcd(I).RY = 0
        End If
        Randomize
        Paramcd(I).RZ = Num2 * Fix(360 * Rnd) + 1
        Randomize
        Paramcd(I).X = Num2 * Fix(60 * Rnd) + 1
        Randomize
        Paramcd(I).Y = Num2 * Fix(30 * Rnd) + 1
        CD(I).SetPosition Paramcd(I).X, Paramcd(I).Y, 50
        CD(I).SetRotation 0, Paramcd(I).RY, Paramcd(I).RZ
    Next
    
    
    If InfoUtente.Accesso = utente Then
        For I = 1 To NCD - 1
            For K = 1 To IPref
                If Paramcd(I).ID = CDPreferiti(K) Then
                    CD(I).enable True
                    TV = True
                End If
            Next
            If TV = False Then
                CD(I).enable False
            End If
        Next
    End If
    'Carico l'immagine delle due frecce laterali
    TextureFAC.LoadTexture App.Path & "\Arrow.bmp", "Freccia", , , TV_COLORKEY_BLACK
    TextureFAC.LoadTexture App.Path & "\Screen.jpg", "Screen", , , TV_COLORKEY_BLACK
    TextureFAC.LoadTexture App.Path & "\Tasto2.jpg", "Tasto2", , , TV_COLORKEY_BLACK
    TextureFAC.LoadTexture App.Path & "\Tasto5.jpg", "Tasto5", , , TV_COLORKEY_BLACK
    Scena.SetSceneBackGround 0, 0, 0.3
End Sub

Private Sub Prendi_CD(CDNumber As Integer)
    'Dichiaro una variabile che servirà al programma per capire se le posizioni
    'del cd da prelevare sono state settate correttamente
    Dim POSOK As Boolean
   If POSOK = False Then
        If Pos(CDNumber).X < 0 Then
            Pos(CDNumber).X = Pos(CDNumber).X + 0.2
        ElseIf Pos(CDNumber).X > 0 Then
            Pos(CDNumber).X = Pos(CDNumber).X - 0.2
        End If
        If Pos(CDNumber).Y < 0 Then
            Pos(CDNumber).Y = Pos(CDNumber).Y + 0.2
        ElseIf Pos(CDNumber).Y > 0 Then
            Pos(CDNumber).Y = Pos(CDNumber).Y - 0.2
        End If
        If Pos(CDNumber).Z > 20 Then
            Pos(CDNumber).Z = Pos(CDNumber).Z - 0.2
        End If
    End If
    If Pos(CDNumber).RZ < 0 Then
        Pos(CDNumber).RZ = Pos(CDNumber).RZ + 0.5
    ElseIf Pos(CDNumber).RZ > 0 Then
        Pos(CDNumber).RZ = Pos(CDNumber).RZ - 0.5
    End If
    If Pos(CDNumber).RY < 180 Then
        Pos(CDNumber).RY = Pos(CDNumber).RY + 0.5
    ElseIf Pos(CDNumber).RY > 180 Then
        Pos(CDNumber).RY = Pos(CDNumber).RY - 0.5
    End If
        
    If Pos(CDNumber).X > -0.2 And Pos(CDNumber).X < 0.2 And Pos(CDNumber).Y > -0.2 And Pos(CDNumber).Y < 0.2 And Pos(CDNumber).Z > -19.8 And Pos(CDNumber).Z < 20.2 Then
        POSOK = True
    Else
        CD(CDNumber).SetPosition Pos(CDNumber).X, Pos(CDNumber).Y, Pos(CDNumber).Z
    End If
    If Pos(CDNumber).RZ > -0.2 And Pos(CDNumber).RZ < 0.2 And Pos(CDNumber).RY > 179.8 And Pos(CDNumber).RY < 180.2 And POSOK = True Then
        Preso = True
    Else
        CD(CDNumber).SetRotation Pos(CDNumber).RX, Pos(CDNumber).RY, Pos(CDNumber).RZ
        'Reinizializzo la variabile Rotazione
        Rotazione = 0
    End If
    
    'Schermo.DrawText "Prendi " + CStr(CDNumber), 80, 200, RGBA(1, 1, 1, 1)
End Sub

Private Sub Posa_CD(CDNumber As Integer)

    'Dichiaro una variabile che servirà al programma per capire se le posizioni
    'del cd da prelevare sono state settate correttamente
    Dim POSOK As Boolean
    Dim RotOK As Boolean
   If POSOK = False Then
        If Pos(CDNumber).X < Paramcd(CDNumber).X Then
            Pos(CDNumber).X = Pos(CDNumber).X + 0.2
        ElseIf Pos(CDNumber).X > Paramcd(CDNumber).X Then
            Pos(CDNumber).X = Pos(CDNumber).X - 0.2
        End If
        If Pos(CDNumber).Y < Paramcd(CDNumber).Y Then
            Pos(CDNumber).Y = Pos(CDNumber).Y + 0.2
        ElseIf Pos(CDNumber).Y > Paramcd(CDNumber).Y Then
            Pos(CDNumber).Y = Pos(CDNumber).Y - 0.2
        End If
        If Pos(CDNumber).Z < 50 Then
            Pos(CDNumber).Z = Pos(CDNumber).Z + 0.2
        End If
    End If
    If Pos(CDNumber).RZ < Paramcd(CDNumber).RZ Then
        Pos(CDNumber).RZ = Pos(CDNumber).RZ + 0.5
    ElseIf Pos(CDNumber).RZ > Paramcd(CDNumber).RZ Then
        Pos(CDNumber).RZ = Pos(CDNumber).RZ - 0.5
    End If
    If Pos(CDNumber).RY < Paramcd(CDNumber).RY Then
        Pos(CDNumber).RY = Pos(CDNumber).RY + 0.5
    ElseIf Pos(CDNumber).RY > Paramcd(CDNumber).RY Then
        Pos(CDNumber).RY = Pos(CDNumber).RY - 0.5
    End If
        
    If Pos(CDNumber).X > Paramcd(CDNumber).X - 0.2 And Pos(CDNumber).X < Paramcd(CDNumber).X + 0.2 And Pos(CDNumber).Y > Paramcd(CDNumber).Y - 0.2 And Pos(CDNumber).Y < Paramcd(CDNumber).Y + 0.2 And Pos(CDNumber).Z > 48.8 And Pos(CDNumber).Z < 50.2 Then
        POSOK = True
    Else
        CD(CDNumber).SetPosition Pos(CDNumber).X, Pos(CDNumber).Y, Pos(CDNumber).Z
    End If
    If Pos(CDNumber).RZ > Paramcd(CDNumber).RZ - 0.2 And Pos(CDNumber).RZ < Paramcd(CDNumber).RZ + 0.2 And Pos(CDNumber).RY > Paramcd(CDNumber).RY - 0.2 And Pos(CDNumber).RY < Paramcd(CDNumber).RY + 0.2 Then
        RotOK = True
    Else
        CD(CDNumber).SetRotation Pos(CDNumber).RX, Pos(CDNumber).RY, Pos(CDNumber).RZ
    End If
    If POSOK = True And RotOK = True Then
       Paramcd(CDNumber).Lasciato = True
    End If
    'Schermo.DrawText "Posa " + CStr(CDNumber), 100, 200, RGBA(1, 1, 1, 1)
End Sub

Private Sub Trasferisci_Attributi(K As Integer)
    Pos(K).X = Paramcd(K).X
    Pos(K).Y = Paramcd(K).Y
    Pos(K).Z = 50
    Pos(K).RY = Paramcd(K).RY
    Pos(K).RZ = Paramcd(K).RZ
    FirstCicle = False
End Sub

Public Sub Trova_Canzone(ID_Canzone As String)
    Dim Trovato As Boolean
    Dim I As Integer
    Dim K As Integer
    I = 1
    While I <= NCD And Trovato = False
        For K = 1 To 30
            If RTrim(Paramcd(I).IDCanzone(K)) = ID_Canzone Then
                Buffer(IB) = App.Path & "\MP3\" & RTrim(Paramcd(I).Artista) & "\" & RTrim(Paramcd(I).Album) & "\" & RTrim(Paramcd(I).Canzone(K)) & ".mp3"
                Trovato = True
                IB = IB + 1
            End If
        Next
        I = I + 1
    Wend
        If Trovato = True Then
            
        Else
        
        End If
End Sub

Private Sub Controlla_Buffer()
    Dim I As Integer
    If WM.PlayState <> 2 Then
        IC = IC + 1
        WM.Filename = Buffer(IC)
        On Error Resume Next
        WM.Play
        If Err.Number <> 0 Then
            IC = IC - 1
        End If
        If IC = IB - 1 Then
            For I = 1 To IB
                Buffer(I) = ""
            Next
            IB = 1
            IC = 0
        End If
    Else
        
    End If
    
    'Schermo.DrawText CStr(IB), 10, 150, RGBA(1, 1, 1, 1)
    'Schermo.DrawText CStr(IC), 10, 170, RGBA(1, 1, 1, 1)

End Sub

Private Sub Mittente_Close()
    Mittente.Close
    Computer.Speak "User Disconnected!"
End Sub

Private Sub Mittente_DataArrival(ByVal bytesTotal As Long)
    Dim Messaggio As String
    Mittente.GetData Messaggio
    Riconosci_Messaggio Messaggio
End Sub

Private Sub Timer1_Timer()
    Controlla_Comandi
End Sub

Private Sub Disegna_Menù_Importa()
    Schermo.DrawFilledBox 230, 230, 854, 630, RGBA(0, 0, 0, 1)
    Schermo.DrawFilledBox 200, 200, 824, 600, RGBA(0, 0, 0.4, 1)
   
    
    Schermo.DrawBox 200, 200, 824, 600, RGBA(0, 0, 0, 1)
    Schermo.DrawBox 199, 199, 825, 601, RGBA(0, 0, 0, 1)
    Schermo.DrawFilledBox 201, 201, 824, 229, RGBA(0, 0, 0.6, 1)
    Schermo.DrawLine 199, 230, 825, 230, RGBA(0, 0, 0, 1)
    Schermo.DrawLine 199, 231, 825, 231, RGBA(0, 0, 0, 1)
    Schermo.DrawText "Importa nuovo/i CD", 450, 210, RGBA(1, 1, 1, 1), "Carattere5"
    'Disegno la barra di avanzamento totale vuota
    Schermo.DrawFilledBox 260, 290, 760, 330, RGBA(0, 0, 0, 1)
    'Disegno la barra di avanzamento parziale vuota
    Schermo.DrawFilledBox 260, 390, 760, 430, RGBA(0, 0, 0, 1)
    
    Schermo.DrawLine 260, 280, 450, 280, RGBA(1, 1, 1, 1)
    Schermo.DrawText "Avanzamento totale", 260, 260, RGBA(1, 1, 1, 1), "Carattere5"
    
    Schermo.DrawLine 260, 380, 450, 380, RGBA(1, 1, 1, 1)
    Schermo.DrawText "Avanzamento parziale", 260, 360, RGBA(1, 1, 1, 1), "Carattere5"

    'Disegno l'avanzamento delle due barre progressive
    Schermo.DrawFilledBox 260, 290, 260 + AvanzTOTALE, 330, RGBA(1, 0, 0, 1)
    Schermo.DrawFilledBox 260, 390, 260 + AvanzPARZIALE, 430, RGBA(1, 0, 0, 1)
    
    'Mando a video il valore percentuale di avanzamento totale
    If (AvanzTOTALE / 500) * 100 < 10 Then
        Schermo.DrawText Mid(CStr(AvanzTOTALE / 500) * 100, 1, 1) & " %", 470, 290, RGBA(1, 1, 0, 1), "Carattere1"
    ElseIf (AvanzTOTALE / 500) * 100 < 100 Then
        Schermo.DrawText Mid(CStr(AvanzTOTALE / 500) * 100, 1, 2) & " %", 470, 290, RGBA(1, 1, 0, 1), "Carattere1"
    Else
        Schermo.DrawText Mid(CStr(AvanzTOTALE / 500) * 100, 1, 3) & " %", 470, 290, RGBA(1, 1, 0, 1), "Carattere1"
    End If
    'Mando a video il valore percentuale di avanzamento parziale
    If (AvanzPARZIALE / 500) * 100 < 10 Then
        Schermo.DrawText Mid(CStr(AvanzPARZIALE / 500) * 100, 1, 1) & " %", 470, 390, RGBA(1, 1, 0, 1), "Carattere1"
    ElseIf (AvanzPARZIALE / 500) * 100 < 100 Then
        Schermo.DrawText Mid(CStr(AvanzPARZIALE / 500) * 100, 1, 2) & " %", 470, 390, RGBA(1, 1, 0, 1), "Carattere1"
    Else
        Schermo.DrawText Mid(CStr(AvanzPARZIALE / 500) * 100, 1, 3) & " %", 470, 390, RGBA(1, 1, 0, 1), "Carattere1"
    End If
    
    
    'Schermo.DrawText Mid(CStr(AvanzTOTALE / 500) * 100, 1, 2) & " %", 470, 290, RGBA(1, 1, 0, 1), "Carattere1"
    'Schermo.DrawText Mid(CStr(AvanzPARZIALE / 500) * 100, 1, 2) & " %", 470, 390, RGBA(1, 1, 0, 1), "Carattere1"

    Schermo.DrawText "Aggiunta del brano", 260, 490, RGBA(1, 1, 1, 1), "Carattere5"
    Schermo.DrawLine 260, 510, 450, 510, RGBA(1, 1, 1, 1)

End Sub

Private Sub Copia_Files()
    'Dichiaro un indice temporaneo
    Dim I As Integer
    'Dichiaro un indice per le sottodirectory
    Dim K As Integer
    'Dichiaro un indice per i file
    Dim J As Integer
    'Dichiaro un indice per le traccie
    Dim T As Integer
    'Avvio un ciclo al fine di ricopiare tutte le cartelle all'interno dell'Hard Disk
    'del Jukebox
    For I = 0 To Dir1.ListCount - 1
        'Salvo il nome della directory corrente
        NomeDirectory = Dir1.List(I)
        'In caso ci fosse un errore si prosegue comunque
        On Error Resume Next
        'Creo la nuova directory con lo stesso nome dell'originale (4)
        MkDir "D:\Miei programmi\Jukebox\MP3\" & Mid(NomeDirectory, 7, Len(NomeDirectory))
        'Entro nella cartella appena creata
        Dir1.Path = NomeDirectory
        'Avvio un altro ciclo al fine di creare tutte le sottodirectory presenti all'interno
        'di quella principale
        For K = 0 To Dir1.ListCount - 1
            'Chiamo la funzione addetta al ricavamento delle informazioni del CD
            RicavaInfoCD Mid(Dir1.List(K), 7, Len(Dir1.List(K)))
            'Inizializzo la variabile NumeroTraccia a 1
            NumeroTraccia = 1
            'Creo la nuova sottodirectory con lo stesso nome dell'originale (4)
            MkDir "D:\Miei programmi\Jukebox\MP3\" & Mid(Dir1.List(K), 7, Len(Dir1.List(K)))
            'Se non è avvenuto errore, ovvero la directory non esiste,allora...
            If Err.Number = 0 Or 3022 Then
                'Imposto il Path del controlla File1 all'interno della directory corrente
                File1.Path = Dir1.List(K)
                'Avvio un ciclo al fine di scandire tutti i file presenti all'interno
                'della directory appena creata
                For J = 0 To File1.ListCount - 1
                    DoEvents
                    'Ripulisco lo schermo
                    TV8.Clear
                    'Chiamo la funzione addetta a disegnare il menù di importazione nuovo/i CD
                    Disegna_Menù_Importa
                    Schermo.DrawText Dir1.List(K) & "\" & File1.List(J), 260, 525, RGBA(1, 1, 1, 1), "Carattere6"
                    'Richiamo la funzione addetta al disegno degli elementi grafici del jukebox
                    Disegna_Grafica
                    'Renderizzo tutto su schermo
                    TV8.RenderToScreen
                    'Se il file correntemente esaminato è un MP3, allora...
                    If Mid(File1.List(J), Len(File1.List(J)) - 3, Len(File1.List(J))) = ".mp3" Then
                       'Salvo il nome della traccia privata del suo numero all'interno dell'apposito
                       'array
                        Traccia(CInt(TrackNumber)) = Delete_Track_Number(File1.List(J))
                        'Copio il file correntemente analizzato dal CD nella sua rispettiva directory locata
                       'all'interno dell'Hard Disk del Jukebox
                        FileCopy Dir1.List(K) & "\" & File1.List(J), "D:\Miei programmi\Jukebox\MP3\" & Mid(Dir1.List(K), 7, Len(Dir1.List(K))) & "\" & Traccia(CInt(TrackNumber))
                        'Aggiorno la variabile NumeroTraccia
                        NumeroTraccia = NumeroTraccia + 1
                        'Aggiorno la barra di avanzamento
                        AvanzTOTALE = AvanzTOTALE + (500 / TotFiles)
                        'Aggiorno la seconda barra di avanzamento
                        AvanzPARZIALE = AvanzPARZIALE + (500 / File1.ListCount)
                    'Altrimenti se il file analizzato è un formato .jpg,allora...
                    ElseIf Mid(File1.List(J), Len(File1.List(J)) - 3, Len(File1.List(J))) = ".jpg" Then
                        'Se il file .jpg si chiamo BACK.jpg, allora...
                        If File1.List(J) = "BACK.jpg" Then
                            'Imposto la variabile TextureBack a valore booleano True, in modo
                            'che il programma capirà che è stata trovata la cover del retro del
                            'nuovo CD
                            TextureBack = True
                        'Se il file .jpg si chiamo FRONT.jpg, allora...
                        ElseIf File1.List(J) = "FRONT.jpg" Then
                            'Imposto la variabile TextureFront a valore booleano True, in modo
                            'che il programma capirà che è stata trovata la cover frontale del
                            'nuovo CD
                            TextureFront = True
                            'Richiamo la funzione addetta alla creazione della nuova Texture
                            TmpTexture = Crea_Texture(Dir1.List(K) & "\FRONT.jpg", Dir1.List(K) & "\BACK.jpg")
                        End If
                        'Aggiorno la barra di avanzamento
                        AvanzTOTALE = AvanzTOTALE + (500 / TotFiles)
                        'Aggiorno la seconda barra di avanzamento
                        AvanzPARZIALE = AvanzPARZIALE + (500 / File1.ListCount)
                        'Salvo il nome della texture in una variabile temporanea
                        'TmpTexture = File1.List(J)
                    End If
                'Si passa al file successivo
                Next
                'Imposto a valore massimo la seconda barra di avanzamento
                AvanzPARZIALE = 500
                'Il valore nel rispettivo label
                Schermo.DrawText "100 %", 470, 390, RGBA(1, 1, 0, 1), "Carattere1"
                'Apertura dell'oggetto Recordset
                Set RS = DB.OpenRecordset("CD", dbOpenTable)
                With RS
                    'Aggiunta di un nuovo record
                    RS.AddNew
                    'Definizione dei nuovi campi fields
                    .Fields("ID").Value = 1 + NCD
                    .Fields("Artista").Value = TmpArtista
                    .Fields("Titolo").Value = TmpTitolo
                    .Fields("NumeroTraccie").Value = NumeroTraccia - 1
                    .Fields("Texture").Value = TmpTexture
                    For T = 1 To NumeroTraccia - 1
                        .Fields("Canzone" & CStr(T)).Value = Mid(Traccia(T), 1, Len(Traccia(T)) - 4)
                    Next
                    For T = NumeroTraccia To 30
                        .Fields("Canzone" & CStr(T)).Value = "\"
                    Next
                    'Aggiorna
                    .Update
                End With
                'Aggiorno il numero delle canzoni
                NCanzoni = NCanzoni + NumeroTraccia - 1
                RS.Close
                NCD = NCD + 1
            End If
            'Pulisco l'array traccia
            For T = 1 To NumeroTraccia
                Traccia(T) = ""
            Next
            'Azzero la seconda barra di avanzamento
            AvanzPARZIALE = 0
            'Il valore nel rispettivo label
        'Si passa alla directory successiva
        Next
        'Reimposto la Path alla root del CD
        Dir1.Path = "D:\CD\"
    'Passo alla directory successiva
    Next
    'Infine viene impostato il massimo valore di avanzamento ad entrambe le barre
    AvanzTOTALE = 500
    AvanzPARZIALE = 500
    
    'Restart del JukeBox
    End
End Sub

Private Function Delete_Track_Number(Filename As String) As String
    'Dichiaro un indice temporaneo
    Dim S As Integer
    'Dichiaro una variabile che servirà a contenere i vari caratteri estratti dal
    'FileName
    Dim car As String
    'Dichiaro una variabile booleana che servirà al programma per capire se dovrà
    'incominciare a copiare il titolo della traccia
    Dim InizioTitolo As Boolean
    Dim TmpInizio As Boolean
    TmpInizio = True
    'Azzero il contenuto della variabile TrackNumber
    TrackNumber = ""
    'Eseguo un ciclo di estrazione di tutti i caratteri al fine di ricopiare il titolo
    'della traccia privo di numero
    For S = 1 To Len(Filename)
        'Prelevo un nuovo carattere
        car = Mid(Filename, S, 1)
        'Se la variabile InizioTitolo è stata impostata a True, ovvero si deve copiare il
        'carattere corrente, allora...
        If InizioTitolo = True Then
            If TmpInizio = True Then
                If car <> " " Then
                    'Copio appunto il carattere corrente
                    Delete_Track_Number = Delete_Track_Number & car
                End If
                TmpInizio = False
            Else
                Delete_Track_Number = Delete_Track_Number & car
            End If
        'Altrimenti
        ElseIf InizioTitolo = False And car <> " " And car <> "-" Then
            If TmpInizio = True Then
                If car <> "0" Then
                    TrackNumber = TrackNumber & car
                End If
                TmpInizio = False
            Else
                TrackNumber = TrackNumber & car
            End If
        End If
        'Se il carattere appena prelevato è uguale a -, allora...
        If car = "-" Then
            TmpInizio = True
            'Imposto la variabile InizioTitolo a valore booleano True, in modo che il
            'programma capirà che da qui in poi dovrà copiare tutti i caratteri che verranno
            'man mano prelevati
            InizioTitolo = True
        End If
    'Si passa al carattere successivo
    Next
End Function

Private Sub RicavaInfoCD(Directory As String)
    'Dichiaro un indice temporaneo
    Dim S As Integer
    'Dichiaro una variabile che servirà a contenere i vari caratteri estratti
    Dim car As String
    'Dichiaro una variabile booleana che servirà al programma per capire se dovrà
    'incominciare a copiare il titolo del CD
    Dim InizioTitolo As Boolean
    'Cancello il contenuto delle due variabili temporanee
    TmpArtista = ""
    TmpTitolo = ""
    'Eseguo un ciclo di estrazione di tutti i caratteri al fine di ricopiare il nome dell'artista e
    'il titolo del CD
    For S = 1 To Len(Directory)
        'Prelevo un nuovo carattere
        car = Mid(Directory, S, 1)
        'Se la variabile InizioTitolo è stata impostata a True, ovvero si deve copiare il
        'carattere corrente, allora...
        If InizioTitolo = True Then
            'Copio appunto il carattere corrente all'interno della variabile temporanea che
            'dovrà contenere il titolo del CD
            TmpTitolo = TmpTitolo & car
        'Altrimenti
        ElseIf InizioTitolo = False And car <> "\" Then
            TmpArtista = TmpArtista & car
        End If
        'Se il carattere appena prelevato è uguale a -, allora...
        If car = "\" Then
            'Imposto la variabile InizioTitolo a valore booleano True, in modo che il
            'programma capirà che da qui in poi dovrà copiare tutti i caratteri che verranno
            'man mano prelevati all'interno della variabile temporanea adatta
            InizioTitolo = True
        End If
    'Si passa al carattere successivo
    Next
End Sub

Private Sub Imposta_Barra_Avanzamento()
    'Dichiaro un indice temporaneo
    Dim I As Integer
    'Dichiaro un indice per le sottodirectory
    Dim K As Integer
    'Dichiaro un indice per i file
    Dim J As Integer
    'Avvio un ciclo al fine di ricopiare tutte le cartelle all'interno dell'Hard Disk
    'del Jukebox
    For I = 0 To Dir1.ListCount - 1
        'Salvo il nome della directory corrente
        NomeDirectory = Dir1.List(I)
        'Entro nella cartella appena creata
        Dir1.Path = NomeDirectory
        'Avvio un altro ciclo al fine di creare tutte le sottodirectory presenti all'interno
        'di quella principale
        For K = 0 To Dir1.ListCount - 1
            'Imposto il Path del controlla File1 all'interno della directory corrente
            File1.Path = Dir1.List(K)
            'Avvio un ciclo al fine di scandire tutti i file presenti all'interno
            'della directory appena creata
            For J = 0 To File1.ListCount - 1
                'Incremento il contatore del numero di files presenti
                TotFiles = TotFiles + 1
            'Passo al file successivo
            Next
        'Passo alla Sottodirectory successiva
        Next
        'Reimposto la Path alla root del CD
        Dir1.Path = "D:\CD\"
    'Passo alla directory successiva
    Next
End Sub

Private Function Crea_Texture(CoverFront As String, CoverBack As String) As String
    With FormTexture
        .Texture.Picture = LoadPicture("D:\Miei programmi\Jukebox\3D Studio Max - Photoshop Files\CDCoverLayout.jpg")
        .BarraD.Picture = LoadPicture("")
        .BarraS.Picture = LoadPicture("")
        .Front.Picture = LoadPicture("")
        .Back.Picture = LoadPicture("")
        '------------------------------------------------------------------
        ' Copia della Cover frontale sul layout della Texture del nuovo CD
        '------------------------------------------------------------------
        .Front.Picture = LoadPicture(CoverFront)
        .Texture.PaintPicture .Front.Picture, 510, 4180, .Front.Width - 40, .Front.Height - 30
        .Texture.Picture = .Texture.Image
        '------------------------------------------------------------------
        ' Copia della Cover del retro sul layout della Texture del nuovo CD
        '------------------------------------------------------------------
        .Back.Picture = LoadPicture(CoverBack)
        .Texture.PaintPicture .Back.Picture, 3705, 4200, .Back.Width - 250, .Back.Height - 40, 800
        .Texture.Picture = .Texture.Image
    
        .BarraD.PaintPicture .Back.Picture, -6060, 1, .Back.Width + 2300, .Back.Height + 20
        .BarraD.Picture = .BarraD.Image
        .Texture.PaintPicture .BarraD.Picture, 4060, 900, .BarraD.Width, .BarraD.Height + 20
        .Texture.Picture = .Texture.Image
    
        .BarraS.PaintPicture .Back.Picture, 1, 1, .Back.Width + 2000, .Back.Height
        .BarraS.Picture = .BarraS.Image
        .Texture.PaintPicture .BarraS.Picture, 3060, 900, .BarraS.Width, .BarraS.Height
        .Texture.Picture = .Texture.Image
        '------------------------------------------------------------------
        ' Salvataggio nuova texture
        '------------------------------------------------------------------
        Crea_Texture = "\CD COVERS\" & TmpArtista & " - " & TmpTitolo & ".jpg"
        SavePicture .Texture.Picture, App.Path & "\CD COVERS\" & TmpArtista & " - " & TmpTitolo & ".jpg"
        TextureFront = False
        TextureBack = False
    End With
End Function

Private Sub LoadingCD()
    Dim I As Integer
    Dim ValTesto As String
    Loading.Play
    For I = 1 To NCD
        ValTesto = ValTesto & RTrim(Paramcd(I).Artista) & " - " & RTrim(Paramcd(I).Album) & "#"
    Next
    Loading.SetVariable "Testo", ValTesto
    Loading.SetVariable "Lunghezza", Len(ValTesto)
End Sub

Private Sub Loading_FSCommand(ByVal command As String, ByVal args As String)
    OKFLASH = True
End Sub

Private Sub Attiva_CD()
    'Dichiarazione di una variabile temporanea
    Dim I As Integer
    'Viene eseguito un ciclo addetta all'abilitazione di tutti i cd
    For I = 1 To NCD - 1
        'Abilito il cd corrente
        CD(I).enable True
    'Si passa al CD successivo
    Next
End Sub

Private Sub Distruggi()
    Dim I As Integer
    'Set TV8 = Nothing
    'Set Scena = Nothing
    'Set Schermo = Nothing
    'Set Effetti = Nothing
    'Set TextureFAC = Nothing
    'Set Comandi = Nothing
    For I = 1 To NCD + 1
        Set CD(I) = Nothing
    Next
    
End Sub
