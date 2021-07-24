Attribute VB_Name = "General"
'AoYind 3.0.0
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Global LeerNPCs As New clsIniReader

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/14/07
'Da cuerpo desnudo a un usuario
'***************************************************
Dim CuerpoDesnudo As Integer
Select Case UserList(UserIndex).genero
    Case eGenero.Hombre
        Select Case UserList(UserIndex).raza
            Case eRaza.Humano
                CuerpoDesnudo = 21
            Case eRaza.Drow
                CuerpoDesnudo = 32
            Case eRaza.Elfo
                CuerpoDesnudo = 210
            Case eRaza.Gnomo
                CuerpoDesnudo = 222
            Case eRaza.Enano
                CuerpoDesnudo = 53
        End Select
    Case eGenero.Mujer
        Select Case UserList(UserIndex).raza
            Case eRaza.Humano
                CuerpoDesnudo = 39
            Case eRaza.Drow
                CuerpoDesnudo = 40
            Case eRaza.Elfo
                CuerpoDesnudo = 259
            Case eRaza.Gnomo
                CuerpoDesnudo = 260
            Case eRaza.Enano
                CuerpoDesnudo = 60
        End Select
End Select

If Mimetizado Then
    UserList(UserIndex).CharMimetizado.Body = CuerpoDesnudo
Else
    UserList(UserIndex).Char.Body = CuerpoDesnudo
End If

UserList(UserIndex).flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal B As Boolean)
'b ahora es boolean,
'b=true bloquea el tile en (x,y)
'b=false desbloquea el tile en (x,y)
'toMap = true -> Envia los datos a todo el mapa
'toMap = false -> Envia los datos al user
'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s

If toMap Then
    Call SendToAreaByPos(UserList(sndIndex).Pos.map, X, Y, PrepareMessageBlockPosition(X, Y, B))
Else
    Call WriteBlockPosition(sndIndex, X, Y, B)
End If

End Sub


Function HayAgua(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

If map > 0 And map < NumMaps + 1 And X > 0 And X <= MapInfo(map).Width And Y > 0 And Y <= MapInfo(map).Height Then
    If ((MapData(map).Tile(X, Y).Graphic(1) >= 1505 And MapData(map).Tile(X, Y).Graphic(1) <= 1520) Or _
    (MapData(map).Tile(X, Y).Graphic(1) >= 5665 And MapData(map).Tile(X, Y).Graphic(1) <= 5680) Or _
    (MapData(map).Tile(X, Y).Graphic(1) >= 13547 And MapData(map).Tile(X, Y).Graphic(1) <= 13562)) And _
       MapData(map).Tile(X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function

Private Function HayLava(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'***************************************************
If map > 0 And map < NumMaps + 1 And X > 0 And X <= MapInfo(map).Width And Y > 0 And Y <= MapInfo(map).Height Then
    If MapData(map).Tile(X, Y).Graphic(1) >= 5837 And MapData(map).Tile(X, Y).Graphic(1) <= 5852 Then
        HayLava = True
    Else
        HayLava = False
    End If
Else
  HayLava = False
End If

End Function


Sub LimpiarMundo()
'***************************************************
'Author: Unknow
'Last Modification: 04/15/2008
'01/14/2008: Marcos Martinez (ByVal) - La funcion FOR estaba mal. En ves de i habia un 1.
'04/15/2008: (NicoNZ) - La funcion FOR estaba mal, de la forma que se hacia tiraba error.
'***************************************************
On Error GoTo errhandler

Dim i As Integer
Dim d As New cGarbage

For i = TrashCollector.Count To 1 Step -1
    Set d = TrashCollector(i)
    Call EraseObj(1, d.map, d.X, d.Y)
    Call TrashCollector.Remove(i)
    Set d = Nothing
Next i

Call SecurityIp.IpSecurityMantenimientoLista

Exit Sub

errhandler:
    Call LogError("Error producido en el sub LimpiarMundo: " & Err.Description)
End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
Dim k As Long
Dim npcNames() As String

ReDim npcNames(1 To UBound(SpawnList)) As String

For k = 1 To UBound(SpawnList)
    npcNames(k) = SpawnList(k).NpcName
Next k

Call WriteSpawnList(UserIndex, npcNames())

End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal port As Integer)
#If UsarQueSocket = 0 Then

Obj.AddressFamily = AF_INET
Obj.Protocol = IPPROTO_IP
Obj.SocketType = SOCK_STREAM
Obj.Binary = False
Obj.Blocking = False
Obj.BufferSize = 1024
Obj.LocalPort = port
Obj.backlog = 5
Obj.listen

#End If
End Sub




Sub Main()
    On Error Resume Next
    Dim f As Date

    ChDir App.Path
    ChDrive App.Path
    'eventos
    Call LoadMapEvent
    'eventos
    'bots
    'inicializa los spells de los bots
    Call ModBots.ia_Spells
    'bots
    Call LoadMotd
    Call BanIpCargar

    Prision.map = 1
    Libertad.map = 1
    Call LoadRanking
    Prision.X = 240
    Prision.Y = 225
    Libertad.X = 240
    Libertad.Y = 240


    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")

    IniPath = App.Path & "\"

    centinelaActivado = True

    LevelSkill(1).LevelValue = 3
    LevelSkill(2).LevelValue = 5
    LevelSkill(3).LevelValue = 7
    LevelSkill(4).LevelValue = 10
    LevelSkill(5).LevelValue = 13
    LevelSkill(6).LevelValue = 15
    LevelSkill(7).LevelValue = 17
    LevelSkill(8).LevelValue = 20
    LevelSkill(9).LevelValue = 23
    LevelSkill(10).LevelValue = 25
    LevelSkill(11).LevelValue = 27
    LevelSkill(12).LevelValue = 30
    LevelSkill(13).LevelValue = 33
    LevelSkill(14).LevelValue = 35
    LevelSkill(15).LevelValue = 37
    LevelSkill(16).LevelValue = 40
    LevelSkill(17).LevelValue = 43
    LevelSkill(18).LevelValue = 45
    LevelSkill(19).LevelValue = 47
    LevelSkill(20).LevelValue = 50
    LevelSkill(21).LevelValue = 53
    LevelSkill(22).LevelValue = 55
    LevelSkill(23).LevelValue = 57
    LevelSkill(24).LevelValue = 60
    LevelSkill(25).LevelValue = 63
    LevelSkill(26).LevelValue = 65
    LevelSkill(27).LevelValue = 67
    LevelSkill(28).LevelValue = 70
    LevelSkill(29).LevelValue = 73
    LevelSkill(30).LevelValue = 75
    LevelSkill(31).LevelValue = 77
    LevelSkill(32).LevelValue = 80
    LevelSkill(33).LevelValue = 83
    LevelSkill(34).LevelValue = 85
    LevelSkill(35).LevelValue = 87
    LevelSkill(36).LevelValue = 90
    LevelSkill(37).LevelValue = 93
    LevelSkill(38).LevelValue = 95
    LevelSkill(39).LevelValue = 97
    LevelSkill(40).LevelValue = 100
    LevelSkill(41).LevelValue = 100
    LevelSkill(42).LevelValue = 100
    LevelSkill(43).LevelValue = 100
    LevelSkill(44).LevelValue = 100
    LevelSkill(45).LevelValue = 100
    LevelSkill(46).LevelValue = 100
    LevelSkill(47).LevelValue = 100
    LevelSkill(48).LevelValue = 100
    LevelSkill(49).LevelValue = 100
    LevelSkill(50).LevelValue = 100


    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.Drow) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.PALADIN) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"

    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasi�n en combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apu�alar) = "Apu�alar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar arboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"

    ListaAtributos(eAtributos.fuerza) = "Fuerza"
    ListaAtributos(eAtributos.agilidad) = "Agilidad"
    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
    ListaAtributos(eAtributos.Carisma) = "Carisma"
    ListaAtributos(eAtributos.Constitucion) = "Constitucion"


    frmCargando.Show

    'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")

    frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    IniPath = App.Path & "\"
    CharPath = App.Path & "\Charfile\"


    DoEvents

    frmCargando.Label1(2).Caption = "Iniciando Arrays..."



    Call CargarSpawnList
    Call CargarForbidenWords
    '�?�?�?�?�?�?�?� CARGAMOS DATOS DESDE ARCHIVOS �??�?�?�?�?�?�?�
    frmCargando.Label1(2).Caption = "Cargando Server.ini"

    MaxUsers = 0
    Call LoadSini
    Call CargaApuestas
    frmCargando.Label1(2).Caption = "Cargando Sets.Dat"
    Call CargarSets
    '*************************************************
    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
    Call CargaNpcsDat
    '*************************************************

    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    'Call LoadOBJData
    Call LoadOBJData

    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call CargarHechizos

    'Jota y a�adido por Neo
    frmCargando.Label1(2).Caption = "Cargando conf. de trabajadores"
    Call CargarConfiguracionTrabajador
    frmCargando.Label1(2).Caption = "Cargando NpcInteligente.Dat"
    Call CargarNpcInteligente
    frmCargando.Label1(2).Caption = "Cargando rutas de carpinteros"
    Call CargarRutaCarpinteros
    frmCargando.Label1(2).Caption = "Cargando rutas de mineros"
    Call CargarRutaMineros
    frmCargando.Label1(2).Caption = "Cargando rutas de herreros"
    Call CargarRutHerreros
    frmCargando.Label1(2).Caption = "Cargando rutas de pescadores"
    Call CargarRutaPescador
    frmCargando.Label1(2).Caption = "Cargando rutas de taladores"
    Call CargarRutaTalador
    frmCargando.Label1(2).Caption = "Cargando frases de trabajadores"
    Call CargarFrasesTrabajador
    frmCargando.Label1(2).Caption = "Cargando trabajadores"
    Call CargarNpcTrabajador
    'Jota y a�adido por Neo
    frmCargando.Label1(2).Caption = "Cargando Objetos de Herrer�a"
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria

    frmCargando.Label1(2).Caption = "Cargando Objetos de Carpinter�a"
    Call LoadObjCarpintero

    frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
    Call LoadBalance    '4/01/08 Pablo ToxicWaste

    frmCargando.Label1(2).Caption = "Cargando Zonas"
    Call CargarZonas

    frmCargando.Label1(2).Caption = "Conectando a MySQL"
    Set mySQL = New clsMySQL
    If mySQL.SQLConnect(MySQL_Host, MySQL_DB, MySQL_User, MySQL_Pass) Then
        frmCargando.Label1(2).Caption = "Conexi�n Exitosa"
    Else
        MsgBox "No se ha podido conectar a la base de datos.", vbCritical
        End
    End If

    frmCargando.Label1(2).Caption = "Cargando Mapas"
    Call LoadMapData

    If BootDelBackUp Then
        frmCargando.Label1(2).Caption = "Cargando BackUp"
        Call CargarBackUp
    End If

    DoEvents
    frmCargando.Label1(2).Caption = "Generando Areas..."
    Call CargarAreas
    frmCargando.Label1(2).Caption = "Finalizando..."
    'Comentado porque hay worldsave en ese mapa!
    Call CrearClanPretoriano(Zonas(MAPA_PRETORIANO).X1)
    '�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�

    Dim LoopC As Integer

    'Resetea las conexiones de los usuarios
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC

    '�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�

    With frmMain
        .AutoSave.Enabled = True
        .tPiqueteC.Enabled = True
        .GameTimer.Enabled = True
        .tLluviaEvent.Enabled = True
        .Auditoria.Enabled = True
        .KillLog.Enabled = True
        .TIMER_AI.Enabled = True
        .npcataca.Enabled = True

        #If SeguridadAlkon Then
            .securityTimer.Enabled = True
        #End If
    End With

    '�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�
    'Configuracion de los sockets

    Call SecurityIp.InitIpTables(1000)

    #If UsarQueSocket = 1 Then

        Call IniciaWsApi(frmMain.hWnd)
        SockListen = ListenForConnect(Puerto, hWndMsg, "")

    #ElseIf UsarQueSocket = 0 Then

        frmCargando.Label1(2).Caption = "Configurando Sockets"

        frmMain.Socket2(0).AddressFamily = AF_INET
        frmMain.Socket2(0).Protocol = IPPROTO_IP
        frmMain.Socket2(0).SocketType = SOCK_STREAM
        frmMain.Socket2(0).Binary = False
        frmMain.Socket2(0).Blocking = False
        frmMain.Socket2(0).BufferSize = 2048

        Call ConfigListeningSocket(frmMain.Socket1, Puerto)

    #ElseIf UsarQueSocket = 2 Then

        frmMain.Serv.Iniciar Puerto

    #ElseIf UsarQueSocket = 3 Then

        frmMain.TCPServ.Encolar True
        frmMain.TCPServ.IniciarTabla 1009
        frmMain.TCPServ.SetQueueLim 51200
        frmMain.TCPServ.Iniciar Puerto

    #End If

    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    '�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�




    Unload frmCargando

    Call LoadGuildsDB

    'quest
    Call LoadQuests
    'quest

    'Jota y a�adido por Neo
    Call initMercader
    'Jota y a�adido por Neo
    Call initFortalezas


    'Log
    Dim N As Integer
    N = FreeFile
    Open LogPath & "\Main.log" For Append Shared As #N
    Print #N, Date & " " & time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
    Close #N

    'Ocultar
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If


    tInicioServer = (GetTickCount() And &H7FFFFFFF)
    Call InicializaEstadisticas
    Call InitBarcos
End Sub

Function FileExist(ByVal file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
    FileExist = LenB(Dir$(file, FileType)) <> 0
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a string
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    Dim Delimiter As String * 1
    
    Delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, Delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function

Function MapaValido(ByVal map As Integer) As Boolean
MapaValido = map >= 1 And map <= NumMaps
End Function

Sub MostrarNumUsers()

frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers

End Sub


Public Sub LogCriticEvent(desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoReal(desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoCaos(desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\EjercitoCaos.log" For Append Shared As #nfile
Print #nfile, desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogIndex(ByVal index As Integer, ByVal desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\" & index & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogError(desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogStatic(desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\Stats.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogTarea(desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile(1) ' obtenemos un canal
Open LogPath & "\haciendo.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & desc
Close #nfile

Exit Sub

errhandler:


End Sub


Public Sub LogClanes(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\clanes.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & str
Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\IP.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & str
Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & str
Close #nfile

End Sub



Public Sub LogGM(Nombre As String, texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
Open LogPath & "\" & Nombre & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogAsesinato(texto As String)
On Error GoTo errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open LogPath & "\asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open LogPath & "\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:


End Sub
Public Sub LogHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogCheating(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\CH.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open LogPath & "\AntiCheat.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Print #nfile, ""
Close #nfile

Exit Sub

errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim arg As String
Dim i As Integer


For i = 1 To 33

arg = ReadField(i, cad, 44)

If LenB(arg) = 0 Then Exit Function

Next i

ValidInputNP = True

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

Dim LoopC As Long
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next

'Initialize statistics!!
Call Statistics.Initialize

For LoopC = 1 To UBound(UserList())
    Set UserList(LoopC).incomingData = Nothing
    Set UserList(LoopC).outgoingData = Nothing
Next LoopC

ReDim UserList(1 To MaxUsers) As User

For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
    Set UserList(LoopC).incomingData = New clsByteQueue
    Set UserList(LoopC).outgoingData = New clsByteQueue
Next LoopC

LastUser = 0
NumUsers = 0

Call FreeNPCs
Call FreeCharIndexes

Call LoadSini
Call LoadOBJData

Call LoadMapData

Call CargarHechizos

#If UsarQueSocket = 0 Then

'*****************Setup socket
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = val(Puerto)
frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

'Log it
Dim N As Integer
N = FreeFile
Open LogPath & "\Main.log" For Append Shared As #N
Print #N, Date & " " & time & " servidor reiniciado."
Close #N

'Ocultar

If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

  
End Sub




Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
    
    If Zonas(UserList(UserIndex).zona).Terreno <> eTerreno.Dungeon Then
        If MapData(UserList(UserIndex).Pos.map).Tile(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger <> 1 And _
           MapData(UserList(UserIndex).Pos.map).Tile(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger <> 2 And _
           MapData(UserList(UserIndex).Pos.map).Tile(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
    
End Function

Public Sub EfectoLluvia(ByVal UserIndex As Integer)
On Error GoTo errhandler


If UserList(UserIndex).flags.UserLogged Then
    If Intemperie(UserIndex) Then
                Dim modifi As Long
                modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 3)
                Call QuitarSta(UserIndex, modifi)
                Call FlushBuffer(UserIndex)
    End If
End If

Exit Sub
errhandler:
 LogError ("Error en EfectoLluvia")
End Sub


Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
    
    Dim modifi As Integer
    
    If UserList(UserIndex).Counters.Frio < IntervaloFrio Then
        UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + 1
    Else
        If Zonas(UserList(UserIndex).zona).Terreno = eTerreno.Nieve Then
            Call WriteConsoleMsg(UserIndex, "��Estas muriendo de frio, abrigate o moriras!!.", FontTypeNames.FONTTYPE_INFO)
            modifi = Porcentaje(UserList(UserIndex).Stats.MaxHP, 5)
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi
            
            If UserList(UserIndex).Stats.MinHP < 1 Then
                Call WriteConsoleMsg(UserIndex, "��Has muerto de frio!!.", FontTypeNames.FONTTYPE_INFO)
                UserList(UserIndex).Stats.MinHP = 0
                Call UserDie(UserIndex)
            End If
            
            Call WriteUpdateHP(UserIndex)
        End If
        
        UserList(UserIndex).Counters.Frio = 0
    End If
End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'If user is standing on lava, take health points from him
'***************************************************
    If UserList(UserIndex).Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
        UserList(UserIndex).Counters.Lava = UserList(UserIndex).Counters.Lava + 1
    Else
        If HayLava(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then
            Call WriteConsoleMsg(UserIndex, "��Quitate de la lava, te est�s quemando!!.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Porcentaje(UserList(UserIndex).Stats.MaxHP, 5)
            
            If UserList(UserIndex).Stats.MinHP < 1 Then
                Call WriteConsoleMsg(UserIndex, "��Has muerto quemado!!.", FontTypeNames.FONTTYPE_INFO)
                UserList(UserIndex).Stats.MinHP = 0
                Call UserDie(UserIndex)
            End If
            
            Call WriteUpdateHP(UserIndex)
        End If
        
        UserList(UserIndex).Counters.Lava = 0
    End If
End Sub

'TRIGGERS
Public Sub EfectoSubeHP(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger = TSUBEHP And .Stats.MinHP < .Stats.MaxHP Then
            Dim hp As Integer
            hp = Porcentaje(.Stats.MaxHP, 5)
            .Stats.MinHP = .Stats.MinHP + hp
            If .Stats.MinHP > .Stats.MaxHP Then
                .Stats.MinHP = .Stats.MaxHP
            End If
            
            Call WriteUpdateHP(UserIndex)
            Call WriteTooltip(UserIndex, .Pos.X, .Pos.Y, 6, "+" & hp)
        End If
    End With
End Sub

Public Sub EfectoSubeMana(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger = eTrigger.TSUBEMAN And .Stats.MinMAN < .Stats.MaxMAN Then
            Dim Mana As Integer
            Mana = Porcentaje(.Stats.MaxMAN, 5)
            .Stats.MinMAN = .Stats.MinMAN + Mana
            If .Stats.MinMAN > .Stats.MaxMAN Then
                .Stats.MinMAN = .Stats.MaxMAN
            End If
            
            Call WriteUpdateMana(UserIndex)
            Call WriteTooltip(UserIndex, .Pos.X, .Pos.Y, 7, "+" & Mana)
        End If
    End With
End Sub

Public Sub EfectoSubeSTA(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger = eTrigger.tSUBESTAMINA And .Stats.MinSta < .Stats.MaxSta Then
            Dim sta As Integer
            sta = Porcentaje(.Stats.MaxSta, 5)
            .Stats.MinSta = .Stats.MinSta + sta
            If .Stats.MinSta > .Stats.MaxSta Then
                .Stats.MinSta = .Stats.MaxSta
            End If
            
            Call WriteUpdateSta(UserIndex)
            Call WriteTooltip(UserIndex, .Pos.X, .Pos.Y, 8, "+" & sta)
        End If
    End With
End Sub

Public Sub EfectoSubeFuerza(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger = eTrigger.TSUBEFUERZA And .Stats.UserAtributos(eAtributos.fuerza) < MAXATRIBUTOS Then
            Dim fuerza As Integer
            fuerza = Porcentaje(MAXATRIBUTOS, 2)
            .Stats.UserAtributos(eAtributos.fuerza) = .Stats.UserAtributos(eAtributos.fuerza) + fuerza
            If .Stats.UserAtributos(eAtributos.fuerza) > MAXATRIBUTOS Then
                .Stats.UserAtributos(eAtributos.fuerza) = MAXATRIBUTOS
            End If
            
            Call WriteAttributes(UserIndex, True)
            Call WriteTooltip(UserIndex, .Pos.X, .Pos.Y, 9, "+" & fuerza)
            .flags.DuracionEfecto = 700
        End If
    End With
End Sub

Public Sub EfectoSubeAgilidad(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger = eTrigger.TSUBEAGILIDAD And .Stats.UserAtributos(eAtributos.agilidad) < MAXATRIBUTOS Then
            Dim agilidad As Integer
            agilidad = Porcentaje(MAXATRIBUTOS, 2)
            .Stats.UserAtributos(eAtributos.agilidad) = .Stats.UserAtributos(eAtributos.agilidad) + agilidad
            If .Stats.UserAtributos(eAtributos.agilidad) > MAXATRIBUTOS Then
                .Stats.UserAtributos(eAtributos.agilidad) = MAXATRIBUTOS
            End If
            
            Call WriteAttributes(UserIndex, True)
            Call WriteTooltip(UserIndex, .Pos.X, .Pos.Y, 10, "+" & agilidad)
            .flags.DuracionEfecto = 700
        End If
    End With
End Sub


''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)
'******************************************************
'Author: Unknown
'Last Update: 04/11/2008 (NicoNZ)
'
'******************************************************
    Dim Barco As ObjData
    
    With UserList(UserIndex)
        If .Counters.Mimetismo < IntervaloInvisible Then
            .Counters.Mimetismo = .Counters.Mimetismo + 1
        Else
            'restore old char
            Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
            
            If .flags.Navegando Then
                If .flags.Muerto = 0 Then
                    Call ToggleBoatBody(UserIndex)
                Else
                    .Char.Body = iFragataFantasmal
                End If
                
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
            Else
                .Char.Body = .CharMimetizado.Body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            End If
            
            With .Char
            Call ChangeUserChar(UserIndex, .Body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .alaIndex)
            End With
            
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
        End If
    End With
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/09/2010 (ZaMa)
'16/09/2010: ZaMa - Al perder el invi cuando navegas, no se manda el mensaje de sacar invi (ya estas visible).
'***************************************************

    With UserList(UserIndex)
        If .Counters.Invisibilidad < IntervaloInvisible Then
            .Counters.Invisibilidad = .Counters.Invisibilidad + 1
        Else
            .Counters.Invisibilidad = RandomNumber(-100, 100) ' Invi variable :D
            .flags.invisible = 0
            If .flags.Oculto = 0 Then
                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
                ' Si navega ya esta visible..
                If Not .flags.Navegando = 1 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                End If
                
            End If
        End If
    End With

End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).flags.Paralizado = 0
    Npclist(NpcIndex).flags.Inmovilizado = 0
End If

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Ceguera > 0 Then
    UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
Else
    If UserList(UserIndex).flags.Ceguera = 1 Then
        UserList(UserIndex).flags.Ceguera = 0
        Call WriteBlindNoMore(UserIndex)
    End If
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call WriteDumbNoMore(UserIndex)
    End If

End If


End Sub


Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Paralisis > 0 Then
    UserList(UserIndex).Counters.Paralisis = UserList(UserIndex).Counters.Paralisis - 1
Else
    UserList(UserIndex).flags.Paralizado = 0
    UserList(UserIndex).flags.Inmovilizado = 0
    'UserList(UserIndex).Flags.AdministrativeParalisis = 0
    Call WriteParalizeOK(UserIndex)
End If

End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

With UserList(UserIndex)

    If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger = 1 And _
       MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger = 2 And _
       MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger = 4 Then Exit Sub
    
    Dim massta As Integer
    
    Dim LastSta As Integer
    
    If .flags.Desnudo Or _
    .flags.Nadando = True Or _
    (UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta) Then
         
          If UserList(UserIndex).Counters.STACounter < Intervalo Then
             UserList(UserIndex).Counters.STACounter = UserList(UserIndex).Counters.STACounter + 1
          Else
            LastSta = UserList(UserIndex).Stats.MinSta
          
           
            UserList(UserIndex).Counters.STACounter = 0
           
            massta = RandomNumber(1, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
            If UserList(UserIndex).flags.Nadando = True Then
                UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - massta
                If UserList(UserIndex).Stats.MinSta < 0 Then
                    UserList(UserIndex).Stats.MinSta = 0
                End If
            Else
                If UserList(UserIndex).flags.Desnudo Then massta = massta * 0.6
                If Lloviendo And Intemperie(UserIndex) Then massta = massta * 0.5
            
                .Stats.MinSta = .Stats.MinSta + massta
                If .Stats.MinSta > .Stats.MaxSta Then
                    .Stats.MinSta = .Stats.MaxSta
                End If
            
            End If
            
            If UserList(UserIndex).Stats.MinSta <> LastSta Then EnviarStats = True
        End If
         
    End If
End With

End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)
Dim N As Integer

If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
  UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
Else
  Call WriteConsoleMsg(UserIndex, "Est�s envenenado, si no te curas morir�s.", FontTypeNames.FONTTYPE_VENENO)
  UserList(UserIndex).Counters.Veneno = 0
  N = RandomNumber(1, 5)
  UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - N
  If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
  Call WriteUpdateHP(UserIndex)
End If

End Sub

Public Sub EfectoCongelado(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Congelado < IntervaloParalizado / 2 Then
  UserList(UserIndex).Counters.Congelado = UserList(UserIndex).Counters.Congelado + 1
Else
  Call DescongelarUser(UserIndex)
End If

End Sub

Public Sub EfectoChiquito(ByVal UserIndex As Integer)

If UserList(UserIndex).Counters.Chiquito < IntervaloParalizado * 2 Then
  UserList(UserIndex).Counters.Chiquito = UserList(UserIndex).Counters.Chiquito + 1
Else
  Call Chiquitolin(UserIndex, False)
End If

End Sub

Public Sub EfectoNadando(ByVal UserIndex As Integer)

With UserList(UserIndex)

If .Stats.MinSta = 0 And .Stats.MinHP > 0 Then
  Dim da�o As Byte
  da�o = RandomNumber(5, Porcentaje(.Stats.MaxHP, 8))
  .Stats.MinHP = .Stats.MinHP - da�o
  Call WriteUpdateHP(UserIndex)
  Call WriteTooltip(UserIndex, .Pos.X, .Pos.Y, 0, "-" & da�o)
  If RandomNumber(1, 5) = 1 Then
    Call WriteConsoleMsg(UserIndex, "Est�s exhausto!!", FontTypeNames.FONTTYPE_FIGHT)
  End If
ElseIf UserList(UserIndex).Stats.MinHP < 0 Then
  Call Muere_Embarcacion(UserIndex, True)
End If

End With

End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)

'Controla la duracion de las pociones
If UserList(UserIndex).flags.DuracionEfecto > 0 Then
   UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - 1
   If UserList(UserIndex).flags.DuracionEfecto = 0 Then
        UserList(UserIndex).flags.TomoPocion = False
        UserList(UserIndex).flags.TipoPocion = 0
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
              UserList(UserIndex).Stats.UserAtributos(loopX) = UserList(UserIndex).Stats.UserAtributosBackUP(loopX)
        Next
        Call WriteAttributes(UserIndex, True)
   End If
End If

End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)

If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Exit Sub

'Sed
If UserList(UserIndex).Stats.MinAGU > 0 Then
    If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
        UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
    Else
        UserList(UserIndex).Counters.AGUACounter = 0
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10
        
        If UserList(UserIndex).Stats.MinAGU <= 0 Then
            UserList(UserIndex).Stats.MinAGU = 0
            UserList(UserIndex).flags.Sed = 1
        End If
        
        fenviarAyS = True
    End If
End If

'hambre
If UserList(UserIndex).Stats.MinHam > 0 Then
   If UserList(UserIndex).Counters.COMCounter < IntervaloHambre Then
        UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + 1
   Else
        UserList(UserIndex).Counters.COMCounter = 0
        UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam - 10
        If UserList(UserIndex).Stats.MinHam <= 0 Then
               UserList(UserIndex).Stats.MinHam = 0
               UserList(UserIndex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If

End Sub

Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

If MapData(UserList(UserIndex).Pos.map).Tile(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger = 1 And _
   MapData(UserList(UserIndex).Pos.map).Tile(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger = 2 And _
   MapData(UserList(UserIndex).Pos.map).Tile(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger = 4 Then Exit Sub

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
    If UserList(UserIndex).Counters.HPCounter < Intervalo Then
        UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
    Else
        mashit = RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
        
        UserList(UserIndex).Counters.HPCounter = 0
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + mashit
        If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
        Call WriteConsoleMsg(UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
        EnviarStats = True
    End If
End If

End Sub

Public Sub CargaNpcsDat()
    Dim npcfile As String
    
    npcfile = DatPath & "NPCs.dat"
    Call LeerNPCs.Initialize(npcfile)
End Sub

Sub PasarSegundo()
    On Error GoTo errhandler
    Dim i As Long


    'eventos

    EventosAOyin.LoopEvent
    
    

    If CuentaRegresivaTimer > 0 Then
        If CuentaRegresivaTimer > 1 Then
            Call SendData(SendTarget.ToAll, 0, PrepareCounterMsg(CuentaRegresivaTimer))
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("" & CuentaRegresivaTimer - 1, FontTypeNames.FONTTYPE_GUILD))

        Else
            Call SendData(SendTarget.ToAll, 0, PrepareCounterMsg(CuentaRegresivaTimer))
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("YA!", FontTypeNames.FONTTYPE_FIGHT))

        End If
        CuentaRegresivaTimer = CuentaRegresivaTimer - 1
    End If
    For i = 1 To LastUser
    With UserList(i)
        If .Counters.TimeFight > 0 Then
            .Counters.TimeFight = .Counters.TimeFight - 1

            If .Counters.TimeFight = 0 Then
                WriteConsoleMsg i, "Cuenta� �YA!", FontTypeNames.FONTTYPE_FIGHT

                ' En los duelos desparalizamos el cliente
                If .flags.SlotEvent > 0 Then
                    If Events(.flags.SlotEvent).Modality = Enfrentamientos Then
                        Call WriteUserInEvent(i)
                    End If
                End If


            Else
                WriteConsoleMsg i, "Cuenta� " & .Counters.TimeFight, FontTypeNames.FONTTYPE_GUILD

            End If
        End If

        If .Counters.TimeCastleMode > 0 Then
            .Counters.TimeCastleMode = .Counters.TimeCastleMode - 1

            WriteConsoleMsg i, "Revivir�s en " & .Counters.TimeCastleMode & " ...", FontTypeNames.FONTTYPE_GUILD

            If .Counters.TimeCastleMode <= 0 Then
                EventosAOyin.CastleMode_UserRevive i
                WriteConsoleMsg i, "Has revivido. �Ve a defender a tu Rey!", FontTypeNames.FONTTYPE_GUILD
            End If
        End If

    End With
    'eventos
    

    
        If UserList(i).flags.UserLogged Then
            'Cerrar usuario
            If UserList(i).Counters.Saliendo Then
                UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
                If UserList(i).Counters.Salir <= 0 Then
                    Call WriteConsoleMsg(i, "Gracias por jugar Argentum Online 2020", FontTypeNames.FONTTYPE_INFO)
                    Call WriteDisconnect(i)
                    Call FlushBuffer(i)

                    Call CloseSocket(i)
                End If

            End If

            If UserList(i).flags.Nadando = True And UserList(i).flags.Navegando = 1 Then
                Call EfectoNadando(i)
            End If
        End If
    Next i
    Exit Sub

errhandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.Description & " - " & Err.Number & " - UserIndex: " & i)
    Resume Next
End Sub
 
Public Function ReiniciarAutoUpdate() As Double

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call HacerBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

 
Sub GuardarUsuarios()
    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i)
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    haciendoBK = False
End Sub


Sub InicializaEstadisticas()
Dim Ta As Long
Ta = (GetTickCount() And &H7FFFFFFF)

Call EstadisticasWeb.Inicializa(frmMain.hWnd)
Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)

End Sub

Public Sub FreeNPCs()
'***************************************************
'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all NPC Indexes
'***************************************************
    Dim LoopC As Long
    
    ' Free all NPC indexes
    For LoopC = 1 To MAXNPCS
        Npclist(LoopC).flags.NPCActive = False
    Next LoopC
End Sub

Public Sub FreeCharIndexes()
'***************************************************
'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all char indexes
'***************************************************
    ' Free all char indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
End Sub
Public Sub ReproducirSonido(ByVal Destino As SendTarget, ByVal index As Integer, ByVal SoundIndex As Integer)
    Call SendData(Destino, index, PrepareMessagePlayWave(SoundIndex, UserList(index).Pos.X, UserList(index).Pos.Y))
End Sub
Public Sub TravelingEffect(ByVal UserIndex As Integer)
'******************************************************
'Author: ZaMa
'Last Update: 01/06/2010 (ZaMa)
'******************************************************

    ' Si ya paso el tiempo de penalizacion
    If IntervaloGoHome(UserIndex) Then
        Call HomeArrival(UserIndex)
    End If

End Sub



