VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   6510
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   5385
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6510
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer TimerPreto 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4320
      Top             =   240
   End
   Begin VB.Timer timerbots 
      Interval        =   40
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox pBarcos 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H000000FF&
      Height          =   4500
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   220
      TabIndex        =   8
      Top             =   1800
      Width           =   3300
   End
   Begin VB.Timer tBarcos 
      Interval        =   15
      Left            =   2880
      Top             =   120
   End
   Begin VB.Timer tRetoCheck 
      Interval        =   1000
      Left            =   2400
      Top             =   120
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   480
      Top             =   540
   End
   Begin VB.Timer packetResend 
      Interval        =   10
      Left            =   480
      Top             =   60
   End
   Begin VB.Timer securityTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   960
      Top             =   60
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   1020
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1440
      Top             =   60
   End
   Begin VB.Timer tLluviaEvent 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   960
      Top             =   1020
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   480
      Top             =   1020
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1920
      Top             =   1020
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1920
      Top             =   60
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1935
      Top             =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
      Begin VB.Timer tMovimiento 
         Interval        =   250
         Left            =   840
         Top             =   0
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Broadcast consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Broadcast clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Escuch 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu BOTS 
      Caption         =   "BOTS"
      Begin VB.Menu Agregarbots 
         Caption         =   "Agregar"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AoYind 3.0.0
'Copyright (C) 2002 M?rquez Pablo Ignacio
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
'Calle 3 n?mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C?digo Postal 1900
'Pablo Ignacio M?rquez

Option Explicit

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, Id As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = Id
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
    Dim iUserIndex As Long
    
    For iUserIndex = 1 To MaxUsers
       'Conexion activa? y es un usuario loggeado?
       If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged Then
            'Actualiza el contador de inactividad
            UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
            If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
                Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado..")
                'mato los comercios seguros
                If UserList(iUserIndex).ComUsu.DestUsu > 0 Then
                    If UserList(UserList(iUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        If UserList(UserList(iUserIndex).ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                            Call WriteConsoleMsg(UserList(iUserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                            Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu)
                            Call FlushBuffer(UserList(iUserIndex).ComUsu.DestUsu) 'flush the buffer to send the message right away
                        End If
                    End If
                    Call FinComerciarUsu(iUserIndex)
                End If
                Call Cerrar_Usuario(iUserIndex)
            End If
        End If
    Next iUserIndex
End Sub

Private Sub Agregarbots_Click()
FrmBots.Show
End Sub

Private Sub Auditoria_Timer()
On Error GoTo errhand
Static centinelSecs As Byte

FechaHora = Format(Date, "dd/MM/yyyy HH:nn:ss")

If Hour(Now) <> Hora Then
    Call SendData(ToAll, 0, PrepareChangeHour())
    Hora = Hour(Now)
End If


centinelSecs = centinelSecs + 1

If centinelSecs = 5 Then
    'Every 5 seconds, we try to call the player's attention so it will report the code.
    Call modCentinela.CallUserAttention
    
    centinelSecs = 0
End If

Call PasarSegundo 'sistema de desconexion de 10 segs

Call ActualizaEstadisticasWeb

Exit Sub

errhand:

Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)
Resume Next

End Sub

Private Sub AutoSave_Timer()

On Error GoTo errhandler
'fired every minute
Static Minutos As Long
Static MinutosLatsClean As Long
Static MinsPjesSave As Long

Dim i As Integer
Dim Num As Long

Minutos = Minutos + 1


'Actualizamos el centinela
Call modCentinela.PasarMinutoCentinela

If Minutos = MinutosWs - 1 Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Worldsave en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))
End If

If Minutos >= MinutosWs Then
    Call HacerBackUp
    Call aClon.VaciarColeccion
    Minutos = 0
End If

If MinutosLatsClean >= 15 Then
    MinutosLatsClean = 0
    Call LimpiarMundo
Else
    MinutosLatsClean = MinutosLatsClean + 1
End If

Call PurgarPenas
Call CheckIdleUser
    If FileExist(CarpetaLogs & "\CantidaddeUsuarios.log", vbNormal) Then Kill CarpetaLogs & "\CantidaddeUsuarios.log"
    'Log helios desde Dll 18/08/2021
    Call Escribe(CarpetaLogs & "\CantidaddeUsuarios.log", CStr(NumUsers))


''<<<<<-------- Log the number of users online ------>>>
'Dim N As Integer
'N = FreeFile()
'Open CarpetaLogs & "\numusers.log" For Output Shared As N
'Print #N, NumUsers
'Close #N
''<<<<<-------- Log the number of users online ------>>>

Exit Sub
errhandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.Description)
    Resume Next
End Sub



Private Sub Command1_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

'Save stats!!!
Call Statistics.DumpStatistics

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
Call LimpiaWsApi
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim N As Integer
N = FreeFile
Open CarpetaLogs & "\Main.log" For Append Shared As #N
Print #N, Date & " " & time & " server cerrado."
Close #N

End

End Sub


Private Sub GameTimer_Timer()
    Dim iUserIndex As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS As Boolean
    
On Error GoTo hayerror
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>
   
    
    For iUserIndex = 1 To MaxUsers 'LastUser
        With UserList(iUserIndex)
           'Conexion activa?
           If .ConnID <> -1 Then
                '?User valido?
                
                If .ConnIDValida And .flags.UserLogged Then
                    
                    '[Alejo-18-5]
                    bEnviarStats = False
                    bEnviarAyS = False
                    
                    
                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
                    If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                    
                    If .Counters.Congelado > 0 Then Call EfectoCongelado(iUserIndex)
                    If .Counters.Chiquito > 0 Then Call EfectoChiquito(iUserIndex)
                    
                    'TRIGGERS
                    Call EfectoSubeHP(iUserIndex)
                    Call EfectoSubeMana(iUserIndex)
                    Call EfectoSubeFuerza(iUserIndex)
                    Call EfectoSubeAgilidad(iUserIndex)
                    Call EfectoSubeSTA(iUserIndex)
                    
                    If .flags.Muerto = 0 Then
                        
                        '[Consejeros]
                        If (.flags.Privilegios And PlayerType.user) Then Call EfectoLava(iUserIndex)
                        If .flags.Desnudo <> 0 And (.flags.Privilegios And PlayerType.user) <> 0 Then Call EfectoFrio(iUserIndex)
                        If .flags.Meditando Then Call DoMeditar(iUserIndex)
                        If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.user) <> 0 Then Call EfectoVeneno(iUserIndex)
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex, False)
                        End If
                        If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                        
                        Call DuracionPociones(iUserIndex)
                        
                        Call HambreYSed(iUserIndex, bEnviarAyS)
                        
                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                            If Lloviendo And Intemperie(iUserIndex) Then
                                    If Not .flags.Descansar Then
                                        'No esta descansando
                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloLluvia)
                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                    Else
                                        'esta descansando
                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                        'termina de descansar automaticamente
                                        If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                            Call WriteRestOK(iUserIndex)
                                            Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                            .flags.Descansar = False
                                        End If
                                        
                                    End If
                            Else
                                If Not .flags.Descansar Then
                                'No esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    
                                Else
                                'esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    'termina de descansar automaticamente
                                    If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                        Call WriteRestOK(iUserIndex)
                                        Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                        .flags.Descansar = False
                                    End If
                                    
                                End If
                            End If
                        End If
                        
                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                        If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                    Else
                        If .flags.Traveling <> 0 Then Call TravelingEffect(iUserIndex)
                    End If 'Muerto
                Else 'no esta logeado?
                    'Inactive players will be removed!
                    .Counters.IdleCount = .Counters.IdleCount + 1
                    If .Counters.IdleCount > IntervaloParaConexion Then
                        .Counters.IdleCount = 0
                        Call CloseSocket(iUserIndex)
                    End If
                End If 'UserLogged
                
                'If there is anything to be sent, we send it
                Call FlushBuffer(iUserIndex)
            End If
        End With
    Next iUserIndex
Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)
End Sub

Private Sub mnuCerrar_Click()


If MsgBox("??Atencion!! Si cierra el servidor puede provocar la perdida de datos. ?Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    mySQL.SQLDisconnect
    Dim f
    For Each f In Forms
        Unload f
    Next
End If

End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog_Timer()
On Error Resume Next
'If FileExist(CarpetaLogs & "\connect.log", vbNormal) Then Kill CarpetaLogs & "\connect.log"
If FileExist(CarpetaLogs & "\haciendo.log", vbNormal) Then Kill CarpetaLogs & "\haciendo.log"
If FileExist(CarpetaLogs & "\stats.log", vbNormal) Then Kill CarpetaLogs & "\stats.log"
If FileExist(CarpetaLogs & "\Asesinatos.log", vbNormal) Then Kill CarpetaLogs & "\Asesinatos.log"
If FileExist(CarpetaLogs & "\HackAttemps.log", vbNormal) Then Kill CarpetaLogs & "\HackAttemps.log"
If Not FileExist(CarpetaLogs & "\nokillwsapi.txt") Then
    If FileExist(CarpetaLogs & "\wsapi.log", vbNormal) Then Kill CarpetaLogs & "\wsapi.log"
End If

End Sub

Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim s As String
Dim nid As NOTIFYICONDATA

s = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, s)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub

Private Sub npcataca_Timer()

On Error Resume Next
Dim NPC As Long

For NPC = 1 To LastNPC
    If Npclist(NPC).AttackTimer = 0 Then
        Npclist(NPC).CanAttack = 1
    Else
        Npclist(NPC).AttackTimer = Npclist(NPC).AttackTimer - 1
    End If
Next NPC

End Sub

Private Sub packetResend_Timer()
'***************************************************
'Autor: Juan Mart?n Sotuyo Dodero (Maraxus)
'Last Modification: 04/01/07
'Attempts to resend to the user all data that may be enqueued.
'***************************************************
On Error GoTo errhandler:
    Dim i As Long
     
    For i = 1 To MaxUsers
        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.Length > 0 Then
                Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.Length))
            End If
        End If
    Next i

Exit Sub

errhandler:
    LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.Description)
    Resume Next
End Sub

Private Sub securityTimer_Timer()

#If SeguridadAlkon Then
    Call Security.SecurityCheck
#End If

End Sub

Private Sub SUPERLOG_Click()

End Sub

Private Sub tBarcos_Timer()
    Call CalcularBarcos
End Sub

Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler
Dim NpcIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim UseAI As Integer
Dim mapa As Integer
Dim e_p As Integer

'Barrin 29/9/03
If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
        
        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
                e_p = esPretoriano(NpcIndex)
                If e_p > 0 Then
                    Select Case e_p
                        Case 1  ''clerigo
                            Call PRCLER_AI(NpcIndex)
                        Case 2  ''mago
                            Call PRMAGO_AI(NpcIndex)
                        Case 3  ''cazador
                            Call PRCAZA_AI(NpcIndex)
                        Case 4  ''rey
                            Call PRREY_AI(NpcIndex)
                        Case 5  ''guerre
                            Call PRGUER_AI(NpcIndex)
                    End Select
                Else
                    'Usamos AI si hay algun user en el mapa
                    If Npclist(NpcIndex).flags.Inmovilizado = 1 Or Npclist(NpcIndex).flags.Paralizado = 1 Then
                       Call EfectoParalisisNpc(NpcIndex)
                    End If
                    
                    mapa = Npclist(NpcIndex).Pos.map
                    
                    If mapa > 0 Then
                        If Npclist(NpcIndex).Movement = Personalizado Then
                            Select Case Npclist(NpcIndex).NpcType 'Estos npc se mueven aunque no haya usuarios cerca..
                                Case eNPCType.Mercader
                                    Call MoverMercader(NpcIndex)
                                Case eNPCType.Fortaleza
                                    Call MoverNPCFortaleza(NpcIndex)
                            End Select
                        ElseIf Npclist(NpcIndex).AreasInfo.Users.Count > 0 Then
                            If Npclist(NpcIndex).Movement <> TipoAI.ESTATICO Then
                                Call NPCAI(NpcIndex)
                            End If
                        ElseIf Npclist(NpcIndex).Movement = TipoAI.NpcPathfinding Then
                            Call VolverOrigPos(NpcIndex)
                        End If
                    End If
                End If
        End If
    Next NpcIndex
End If

Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.map & "X: " & Npclist(NpcIndex).Pos.X & "Y: " & Npclist(NpcIndex).Pos.Y)
    Call MuereNpc(NpcIndex, 0)
End Sub


'Bots
Private Sub timerbots_Timer()

'Acci?n de los bots -
    Dim loopX   As Long
    
    For loopX = 1 To MAX_BOTS
    
        If ia_Bot(loopX).Invocado Then ia_Action loopX
    
    Next loopX
    
End Sub
'bots

Private Sub TimerPreto_Timer()
    TiempoPreto = TiempoPreto + 1
    If TiempoPreto = 120 Then

        PretorianosMuerte = 0


        Call CrearClanPretoriano(Zonas(MAPA_PRETORIANO).X1)
        TiempoPreto = 0
        TimerPreto = False
    End If

End Sub

Private Sub tLluviaEvent_Timer()

On Error GoTo ErrorHandler
Static MinutosLloviendo As Long
Static MinutosSinLluvia As Long

If Not Lloviendo Then
    MinutosSinLluvia = MinutosSinLluvia + 1
    If MinutosSinLluvia >= 35 And MinutosSinLluvia < 1440 Then
            If RandomNumber(1, 100) <= 2 Then
                Lloviendo = True
                
                If RandomNumber(1, 3) = 1 Then
                    LloviendoConTormenta = 1
                Else
                    LloviendoConTormenta = 0
                End If
                
                CheckFogatasLluvia
                MinutosSinLluvia = 0
                Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
            End If
    ElseIf MinutosSinLluvia >= 1440 Then
                Lloviendo = True
                
                If RandomNumber(1, 3) = 1 Then
                    LloviendoConTormenta = 1
                Else
                    LloviendoConTormenta = 0
                End If
                
                CheckFogatasLluvia
                MinutosSinLluvia = 0
                Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End If
Else
    MinutosLloviendo = MinutosLloviendo + 1
    If MinutosLloviendo >= 3 Then
            Lloviendo = False
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
            MinutosLloviendo = 0
    Else
            If RandomNumber(1, 100) <= 2 Then
                Lloviendo = False
                MinutosLloviendo = 0
                Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
            End If
    End If
End If

Exit Sub
ErrorHandler:
Call LogError("Error tLluviaTimer")

End Sub

Private Sub tMovimiento_Timer()
On Error GoTo errhandler
Dim i As Integer
    For i = 1 To LastUser
        With UserList(i)
            If .flags.Movimiento > 0 Then
                .flags.Movimiento = .flags.Movimiento - 3
                If .flags.Movimiento < 0 Then .flags.Movimiento = 0
            End If
        End With
        
    Next i
Exit Sub

errhandler:
    Call LogError("Error en tMovimiento " & Err.Number & ": " & Err.Description)
End Sub

Private Sub tPiqueteC_Timer()
    Dim NuevaA As Boolean
    Dim NuevoL As Boolean
    Dim GI As Integer
    Dim Trigger As Byte
    Dim i As Long
    
On Error GoTo errhandler
    For i = 1 To LastUser
        With UserList(i)
            If .flags.UserLogged Then
                Trigger = MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger
                If Trigger = eTrigger.ANTIPIQUETE Or Trigger = eTrigger.MIRARLEFT Or Trigger = eTrigger.MIRARRIGHT Or Trigger = eTrigger.MIRARUP Then
                    .Counters.PiqueteC = .Counters.PiqueteC + 1
                   
                    
                    If .Counters.PiqueteC > 3 Then
                        Call WriteConsoleMsg(i, "?Est?s obstruyendo la v?a p?blica, mu?vete o ser?s encarcelado!", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                    If .Counters.PiqueteC > 15 Then
                        .Counters.PiqueteC = 0
                        Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                    End If
                Else
                    .Counters.PiqueteC = 0
                End If
                
                'ustedes se preguntaran que hace esto aca?
                'bueno la respuesta es simple: el codigo de AO es una mierda y encontrar
                'todos los puntos en los cuales la alineacion puede cambiar es un dolor de
                'huevos, asi que lo controlo aca, cada 6 segundos, lo cual es razonable
        
                GI = .GuildIndex
                If GI > 0 Then
                    NuevaA = False
                    NuevoL = False
                    If Not modGuilds.m_ValidarPermanencia(i, True, NuevaA, NuevoL) Then
                        Call WriteConsoleMsg(i, "Has sido expulsado del clan. ?El clan ha sumado un punto de antifacci?n!", FontTypeNames.FONTTYPE_GUILD)
                    End If
                    If NuevaA Then
                        Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("?El clan ha pasado a tener alineaci?n neutral!", FontTypeNames.FONTTYPE_GUILD))
                        Call LogClanes("El clan cambio de alineacion!")
                    End If
                    If NuevoL Then
                        Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("?El clan tiene un nuevo l?der!", FontTypeNames.FONTTYPE_GUILD))
                        Call LogClanes("El clan tiene nuevo lider!")
                    End If
                End If
                
                Call FlushBuffer(i)
            End If
        End With
    Next i
Exit Sub

errhandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.Description)
End Sub





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal Id As Long)
On Error GoTo errorHandlerNC

    ESCUCHADAS = ESCUCHADAS + 1
    Escuch.Caption = ESCUCHADAS
    
    Dim i As Integer
    
    Dim NewIndex As Integer
    NewIndex = NextOpenUser
    
    If NewIndex <= MaxUsers Then
        'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
        
        TCPServ.SetDato Id, NewIndex
        
        If aDos.MaxConexiones(TCPServ.GetIP(Id)) Then
            Call aDos.RestarConexion(TCPServ.GetIP(Id))
            Call ResetUserSlot(NewIndex)
            Exit Sub
        End If

        If NewIndex > LastUser Then LastUser = NewIndex

        UserList(NewIndex).ConnID = Id
        UserList(NewIndex).ip = TCPServ.GetIP(Id)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = TCPServ.GetIP(Id) Then
                Call ResetUserSlot(NewIndex)
                Exit Sub
            End If
        Next i

    Else
        Call CloseSocket(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If

Exit Sub

errorHandlerNC:
Call LogError("TCPServer::NuevaConexion " & Err.Description)
End Sub

Private Sub TCPServ_Close(ByVal Id As Long, ByVal MiDato As Long)
    On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
    Call CloseSocket(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & Id & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal Id As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
On Error GoTo errorh

With UserList(MiDato)
    Dim Data() As Byte
    Dim Procesado As Boolean
    
    Data = StrConv(Datos, vbFromUnicode)
    'Debug.Print StrConv(.clave, vbUnicode)
    Call DataCorrect(.clave, Data, .iCliente)
    Datos = StrConv(Data, vbUnicode)
    
    Call .incomingData.WriteASCIIStringFixed(Datos)
    
    If .ConnID <> -1 Then
        Procesado = False
        Do While .incomingData.Length And Not Procesado
            Procesado = HandleIncomingData(MiDato)
        Loop
    End If
End With

Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & Id & " error:" & Err.Description)

End Sub

#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub tRetoCheck_Timer()
Dim i As Integer
For i = 1 To NUM_SALAS
    If SalasRetos(i).Segundos > 0 Then
        SalasRetos(i).Segundos = SalasRetos(i).Segundos - 1
        If SalasRetos(i).Segundos = 0 Then
            Call LimpiarSala(i)
        End If
    End If
Next i
'Aprovecho este timer para checkear los respawn de las fortalezas
Call CheckRespawns
End Sub
