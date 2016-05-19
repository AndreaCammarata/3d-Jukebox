Attribute VB_Name = "Funzioni_Connessione"
Option Explicit
'Dichiarazione di una variabile che servirà al programma per capire se
'è arrivata una dedica e dovrà quindi disegnare il menù apposito
Public Dedica As Boolean
'Dichiaro una variabile che servirà a contenere il messaggio della dedica
Public TestoDedica As String
'Dichiaro una variabile che servirà a contenere il numero ID della
'canzone dedicata
Public IDDEdica As String
'Dichiaro una variabile che servirà al programma per capire se il
'mittente della dedica ha ricevuto i deati precedenti ed è pronto a ricevere
'i nuovi.Utile nel processo di sincronizzazione
Public OK As Boolean
'Dichiaro una variabile temporanea che servirà a mantenere la posizione
'del carattere dal quale inizia il testo della dedica
Dim InizioDedica As Integer

'Definizione di una specifica funzione addetta al riconoscimento dei vari
'codici di identificazione dei messaggi provenienti dall'utente remoto
Sub Riconosci_Messaggio(Messaggio As String)
    'Dichiarazione di una variabile codice all'interno della quale andare a
    'salvare il codice di identificazione messaggio estratto dallo stesso
    Dim Codice As String
    'Estrazione del codice dal messaggio
    Codice = Mid(Messaggio, 1, 4)
    'Cancellazione del codice del messaggio dal messaggio stesso
    Messaggio = Mid(Messaggio, 5, Len(Messaggio))
    'Viene verificato il codice del messaggio
    Select Case Codice
    Case "0000":
        OK = True
    'Caso in cui l'utente remoto richiede la lista delle canzoni dedicabili
    Case "0001":
        'Chiamata alla funzione addetta all'invio della lista dedicabili
        Invia_Lista
   'Caso in cui si riceve il messaggio della dedica
    Case "1000":
        'Imposto A True il valore della variabile Dedica in modo che il programma
        'capisca che è in arrivo una dedica
        Dedica = True
        'Richiamo la funzione addetta al'estrazione dal messaggio arrivato, il numero ID
        'della canzone che è stata dedicata
        IDDEdica = Ricava_ID_Dedica(Messaggio)
        'Salvo il testo della dedica
        TestoDedica = Mid(Messaggio, InizioDedica, Len(Messaggio))
        Main.Computer.CurrentMode = 11
        Main.Computer.Speak TestoDedica
        
        Main.Trova_Canzone IDDEdica
    End Select
End Sub

'Definizione di una funzione addetta all'estrazione dal messaggio il numero
'ID della canzone dedicata
Private Function Ricava_ID_Dedica(Messaggio As String) As String
    'Dichiarazione di un carattere temporaneo
    Dim car As String
    'Dichiarazione di un indice temporaneo
    Dim I As Integer
    'Dichiarazione di una variabile di fine ciclo estrazione
    Dim Fine As Boolean
    'Inizializzazione dell'indice carattere
    I = 1
    'Ciclo di estrazione numero canzone
    Do
        'Estrazione di un nuovo carattere
        car = Mid(Messaggio, I, 1)
        'Se il carattere appena estratto è "\",allora
        If car = "\" Then
            'Il numero ID della canzone dedicata è stato estratto
            Fine = True
            'Salvo la posizione del carattere
            InizioDedica = I + 1
        'Altrimenti...
        Else
            'Viene salvato il carattere appena estratto
            Ricava_ID_Dedica = Ricava_ID_Dedica + car
        End If
        'Viene incrementato l'indice di carattere
        I = I + 1
    'Si passa ad esaminare il carattere successivo
    Loop Until Fine = True
End Function

'Definizione di una funzione addetta dell'invio all'utente remoto, della lista
'delle canzoni disponibili
Public Sub Invia_Lista()
    Main.Computer.CurrentMode = 4
    Main.Computer.Speak "CD List requested!"
    'Dichiarazione di un indice temporaneo
    Dim I As Integer
    'Dichiarazione di un secondo indice temporaneo
    Dim K As Integer
    'Invio il numero di CD presenti all'interno del jukebox
    Main.Mittente.SendData "0001" + CStr(NCD)
    'Chiamata alla funzione di sincronizzazione
    Sincronizza
    'Ciclo di scansione di tutti i cd presenti all'interno del jukebox
    For I = 1 To NCD
        With Main.Mittente
            'Invio dell'artista del CD
            .SendData "0001" + RTrim(Paramcd(I).Artista)
            'Chiamata alla funzione di sincronizzazione
            Sincronizza
            'Invio del titolo del CD
            .SendData "0001" + RTrim(Paramcd(I).Album)
            'Chiamata alla funzione di sincronizzazione
            Sincronizza
            'Ciclo di invio della lista di canzoni presenti all'interno del CD
            For K = 1 To 30
                'Invio dell'ID della canzone
                .SendData "0001" + RTrim(Paramcd(I).IDCanzone(K))
                'Chiamata alla funzione di sincronizzazione
                Sincronizza
                'Invio del titolo della canzone
                .SendData "0001" + RTrim(Paramcd(I).Canzone(K))
                'Chiamata alla funzione di sincronizzazione
                Sincronizza
            Next
            'Invio il codice di fine CD
            .SendData "0010"
            'Chiamata alla funzione di sincronizzazione
            Sincronizza
        End With
    Next
    'Viene infine spedito il messaggio di fine invio lista
    Main.Mittente.SendData "1000"
End Sub

Public Sub Inizializza_Connessione()
    With Main
        'Imposto la porta di ascolto
        .Dedica.LocalPort = 313
        'Metto in attesa il JukeBox per una eventuale dedica in arrivo
        .Dedica.Listen
    End With
End Sub

Private Sub Sincronizza()
    '---------------------------------------
    ' Ciclo di sincronizzazione al client
    '---------------------------------------
    Do
        DoEvents
    Loop Until OK = True
    'Riporto a False la variabile OK
    OK = False
End Sub
