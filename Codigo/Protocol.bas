Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 246


Private Enum ClientPacketIDGuild
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildNewWebsite         'NEWWEBSI
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    ShowGuildNews
End Enum

Private Enum ClientPacketIDGM
    GMMessage               '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    onlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    SetCharDescription      '/SETDESC
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    CitizenMessage          '/CIUMSG
    CriminalMessage         '/CRIMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    DumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    createnpc               '/ACC
    CreateNPCWithRespawn    '/RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterPassword           '/APASS
    AlterMail               '/AEMAIL
    AlterName               '/ANAME
    ToggleCentinelActivated '/CENTINELAACTIVADO
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ResetAutoUpdate         '/AUTOUPDATE
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
    AddGM
    CuentaRegresiva         '/CUENTAREGRESIVA TIEMPO
End Enum

Private Enum ServerPacketID
    OpenAccount
    Logged                  ' LOGGED
    ChangeHour
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    UserOfferConfirm
    CancelOfferItem
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    NPCSwing                ' N1
    NPCKillUser             ' 6
    BlockedWithShieldUser   ' 7
    BlockedWithShieldother  ' 8
    UserSwing               ' U1
    SafeModeOn              ' SEGON
    SafeModeOff             ' SEGOFF
    ResuscitationSafeOn
    ResuscitationSafeOff
    NobilityLost            ' PN
    CantUseWhileMeditating  ' M!
    ataca                   ' USER ATACA
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateBankGold
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    NPCHitUser              ' N2
    UserHitNPC              ' U2
    UserAttackedSwing       ' U3
    UserHittedByUser        ' N4
    UserHittedUser          ' N5
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    ShowMessageScroll
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    CreateEfecto
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    Fame                    ' FAMA
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    SetInvisible            ' NOVER
    SetOculto               ' NOVER OCUL
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildMemberInfo
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    Pong
    UpdateTagAndStatus
    UsersOnline
    CommerceChat
    ShowPartyForm
    StopWorking
    RetosAbre
    RetosRespuesta
    
    
    
    'SERVER NEW PACKETS
    SetEquitando            ' EQUITANDO
    SetCongelado            ' CONGELADO
    SetChiquito             ' CHIQUITOLIN
    CreateAreaFX            ' Crear FX en terreno
    PalabrasMagicas         ' Palabras Magiacs
    UserSpellNPC            ' US2
    SetNadando
    
    MultiMessage            'Message in client
    
    FirstInfo               'Primera Informacion
    CuentaRegresiva         'Cuenta
    
    PicInRender             'PicInRender
    
    Quit                    'Efecto Quit
    
    
    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    ShowBarco
    AgregarPasajero
    QuitarPasajero
    QuitarBarco
    goHome
    GotHome
    Tooltip
    
    'quest
    QuestDetails
    QuestListSend
    'quest
    
End Enum

Private Enum ClientPacketID
    BorrarPJ                'BorrarPJ
    OpenAccount             'ALOGIN
    LoginExistingChar       'OLOGIN
    LoginNewChar            'NLOGIN
    LoginNewAccount         'CALOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestFame             'FAMA
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    UserCommerceConfirm
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    MoveSpell               'DESPHE
    MoveBank
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    CommerceChat
    guild
    Online                  '/ONLINE
    Quit                    '/SALIR
    GuildLeave              '/SALIRCLAN
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPAÑAR
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    ComandosVarios                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    UpTime                  '/UPTIME
    PartyLeave              '/SALIRPARTY
    PartyCreate             '/CREARPARTY
    PartyJoin               '/PARTY
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    PartyMessage            '/PMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    PartyOnline             '/ONLINEPARTY
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    bugReport               '/_BUG
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    Punishments             '/PENAS
    ChangePassword          '/CONTRASEÑA
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA ( with parameters )
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    GuildFundate            '/FUNDARCLAN
    PartyKick               '/ECHARPARTY
    PartySetLeader          '/PARTYLIDER
    PartyAcceptMember       '/ACCEPTPARTY
    Ping                    '/PING
    RequestPartyForm
    ItemUpgrade
    InitCrafting
    RetosAbrir              '/RETOS
    RetosCrear
    RetosDecide
    IntercambiarInv
    
    'GM messages
    gm
    WarpMeToTarget          '/TELEPLOC
    Home
    
    
    'CLIENT NEW PACKETS
    Equitar                 'Montar
    DejarMontura            'Dejar Montura
    CreateEfectoClient      'Crear Efecto desde el cliente
    CreateEfectoClientAction     'Hago la accion (daño, etc)
    AnclarEmbarcacion
    'DesAnclarEmbarcacion
    
    'quest
    Quest                   '/QUEST
    QuestAccept
    QuestListRequest
    QuestDetailsRequest
    QuestAbandon
    'quest
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_RETOS
    FONTTYPE_EXP
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
End Enum

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer As New clsByteQueue




''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal Userindex As Integer) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'
'***************************************************
On Error Resume Next
    Dim packetID As Byte
    
    packetID = UserList(Userindex).incomingData.PeekByte()
    
    'Does the packet requires a logged user??
    If Not (packetID = ClientPacketID.LoginExistingChar _
      Or packetID = ClientPacketID.OpenAccount _
      Or packetID = ClientPacketID.LoginNewChar _
      Or packetID = ClientPacketID.LoginNewAccount _
      Or packetID = ClientPacketID.BorrarPJ) Then
        
        'Is the user actually logged?
        If Not UserList(Userindex).flags.UserLogged Then
            Call CloseSocket(Userindex)
            HandleIncomingData = True
            Exit Function
        
        'He is logged. Reset idle counter if id is valid.
        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
            UserList(Userindex).Counters.IdleCount = 0
        End If
    ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
        UserList(Userindex).Counters.IdleCount = 0
        
        'Is the user logged?
        If UserList(Userindex).flags.UserLogged Then
            Call CloseSocket(Userindex)
            HandleIncomingData = True
            Exit Function
        End If
    End If
    
    Select Case packetID
        Case ClientPacketID.BorrarPJ             'ALOGIN
            Call HandleBorrarPJ(Userindex)
            
        Case ClientPacketID.OpenAccount             'ALOGIN
            Call HandleOpenAccount(Userindex)
            
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Call HandleLoginExistingChar(Userindex)
        
        Case ClientPacketID.LoginNewChar            'NLOGIN
            Call HandleLoginNewChar(Userindex)
            
        Case ClientPacketID.LoginNewAccount         'CALOGIN
            Call HandleLoginNewAccount(Userindex)
        
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(Userindex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(Userindex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(Userindex)
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(Userindex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(Userindex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(Userindex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(Userindex)
            
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(Userindex)
        
        Case ClientPacketID.ResuscitationSafeToggle
            Call HandleResuscitationToggle(Userindex)
        
        Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
            Call HandleRequestGuildLeaderInfo(Userindex)
        
        Case ClientPacketID.RequestAtributes        'ATR
            Call HandleRequestAtributes(Userindex)
        
        Case ClientPacketID.RequestFame             'FAMA
            Call HandleRequestFame(Userindex)
        
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(Userindex)
        
        Case ClientPacketID.RequestMiniStats        'FEST
            Call HandleRequestMiniStats(Userindex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(Userindex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(Userindex)
            
        Case ClientPacketID.UserCommerceConfirm
            Call HandleUserCommerceConfirm(Userindex)
            
        Case ClientPacketID.CommerceChat
            Call HandleCommerceChat(Userindex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(Userindex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(Userindex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(Userindex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(Userindex)
        
        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(Userindex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(Userindex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(Userindex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(Userindex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(Userindex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(Userindex)
        
        Case ClientPacketID.CraftBlacksmith         'CNS
            Call HandleCraftBlacksmith(Userindex)
        
        Case ClientPacketID.CraftCarpenter          'CNC
            Call HandleCraftCarpenter(Userindex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(Userindex)
        
        Case ClientPacketID.CreateNewGuild          'CIG
            Call HandleCreateNewGuild(Userindex)
        
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(Userindex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(Userindex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(Userindex)
        
        Case ClientPacketID.ModifySkills            'SKSE
            Call HandleModifySkills(Userindex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(Userindex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(Userindex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(Userindex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(Userindex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(Userindex)
        
        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(Userindex)
            
        Case ClientPacketID.MoveBank
            Call HandleMoveBank(Userindex)
        
        Case ClientPacketID.ClanCodexUpdate         'DESCOD
            Call HandleClanCodexUpdate(Userindex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(Userindex)
        
        Case ClientPacketID.guild
            Call HandleGuild(Userindex)
            
        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(Userindex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(Userindex)
        
        Case ClientPacketID.GuildLeave              '/SALIRCLAN
            Call HandleGuildLeave(Userindex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(Userindex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(Userindex)
        
        Case ClientPacketID.PetFollow               '/ACOMPAÑAR
            Call HandlePetFollow(Userindex)
        
        Case ClientPacketID.TrainList               '/ENTRENAR
            Call HandleTrainList(Userindex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(Userindex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(Userindex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(Userindex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(Userindex)
        
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(Userindex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(Userindex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(Userindex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(Userindex)
        
        Case ClientPacketID.ComandosVarios                  '/ENLISTAR
            Call HandleComandosVarios(Userindex)
        
        Case ClientPacketID.Information             '/INFORMACION
            Call HandleInformation(Userindex)
        
        Case ClientPacketID.Reward                  '/RECOMPENSA
            Call HandleReward(Userindex)
        
        Case ClientPacketID.RequestMOTD             '/MOTD
            Call HandleRequestMOTD(Userindex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(Userindex)
        
        Case ClientPacketID.PartyLeave              '/SALIRPARTY
            Call HandlePartyLeave(Userindex)
        
        Case ClientPacketID.PartyCreate             '/CREARPARTY
            Call HandlePartyCreate(Userindex)
        
        Case ClientPacketID.PartyJoin               '/PARTY
            Call HandlePartyJoin(Userindex)
        
        Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
            Call HandleInquiry(Userindex)
        
        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(Userindex)
        
        Case ClientPacketID.PartyMessage            '/PMSG
            Call HandlePartyMessage(Userindex)
        
        Case ClientPacketID.CentinelReport          '/CENTINELA
            Call HandleCentinelReport(Userindex)
        
        Case ClientPacketID.GuildOnline             '/ONLINECLAN
            Call HandleGuildOnline(Userindex)
        
        Case ClientPacketID.PartyOnline             '/ONLINEPARTY
            Call HandlePartyOnline(Userindex)
        
        Case ClientPacketID.CouncilMessage          '/BMSG
            Call HandleCouncilMessage(Userindex)
        
        Case ClientPacketID.RoleMasterRequest       '/ROL
            Call HandleRoleMasterRequest(Userindex)
        
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(Userindex)
        
        Case ClientPacketID.bugReport               '/_BUG
            Call HandleBugReport(Userindex)
        
        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(Userindex)
        
        Case ClientPacketID.GuildVote               '/VOTO
            Call HandleGuildVote(Userindex)
        
        Case ClientPacketID.Punishments             '/PENAS
            Call HandlePunishments(Userindex)
        
        Case ClientPacketID.ChangePassword          '/CONTRASEÑA
            Call HandleChangePassword(Userindex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(Userindex)
        
        Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
            Call HandleInquiryVote(Userindex)
        
        Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
            Call HandleLeaveFaction(Userindex)
        
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(Userindex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(Userindex)
        
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(Userindex)
        
        Case ClientPacketID.GuildFundate            '/FUNDARCLAN
            Call HandleGuildFundate(Userindex)
        
        Case ClientPacketID.PartyKick               '/ECHARPARTY
            Call HandlePartyKick(Userindex)
        
        Case ClientPacketID.PartySetLeader          '/PARTYLIDER
            Call HandlePartySetLeader(Userindex)
        
        Case ClientPacketID.PartyAcceptMember       '/ACCEPTPARTY
            Call HandlePartyAcceptMember(Userindex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(Userindex)
        
        Case ClientPacketID.InitCrafting
            Call HandleInitCrafting(Userindex)
            
        Case ClientPacketID.ItemUpgrade
            Call HandleItemUpgrade(Userindex)
        
        Case ClientPacketID.RequestPartyForm
            Call HandlePartyForm(Userindex)
            
        Case ClientPacketID.RetosAbrir
            Call HandleRetosAbrir(Userindex)
            
        Case ClientPacketID.RetosCrear
            Call HandleRetosCrear(Userindex)
            
        Case ClientPacketID.RetosDecide
            Call HandleRetosDecide(Userindex)

        Case ClientPacketID.IntercambiarInv
            Call HandleIntercambiarInv(Userindex)
            
        Case ClientPacketID.Home
            Call HandleHome(Userindex)
            
        Case ClientPacketID.Equitar
            Call HandleEquitar(Userindex)
            
        Case ClientPacketID.DejarMontura
            Call HandleDejarMontura(Userindex)
            
        Case ClientPacketID.CreateEfectoClient
            Call HandleCreateEfectoClient(Userindex)
            
        Case ClientPacketID.CreateEfectoClientAction
            Call HandleCreateEfectoClientAction(Userindex)
            
        Case ClientPacketID.AnclarEmbarcacion
            Call HandleAnclarEmbarcacion(Userindex)
            
        'quest
        Case ClientPacketID.Quest                   '/QUEST
            Call HandleQuest(Userindex)
            
        Case ClientPacketID.QuestAccept
            Call HandleQuestAccept(Userindex)
            
        Case ClientPacketID.QuestListRequest
            Call HandleQuestListRequest(Userindex)
            
        Case ClientPacketID.QuestDetailsRequest
            Call HandleQuestDetailsRequest(Userindex)
            
        Case ClientPacketID.QuestAbandon
            Call HandleQuestAbandon(Userindex)
        
        'quest
            
        
        'GM messages
        Case ClientPacketID.gm
            Call HandleGM(Userindex)
        Case ClientPacketID.WarpMeToTarget          '/TELEPLOC
            Call HandleWarpMeToTarget(Userindex)
        #If SeguridadAlkon Then
                Case Else
                    Call HandleIncomingDataEx(Userindex)
        #Else
                Case Else
                    'ERROR : Abort!
                    Call CloseSocket(Userindex)
        #End If
    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(Userindex).incomingData.Length > 0 And Err.Number = 0 Then
        Err.Clear
        Call HandleIncomingData(Userindex)
        HandleIncomingData = False
    ElseIf Err.Number <> 0 And Not Err.Number = UserList(Userindex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.Description & "] " & " Source: " & Err.Source & _
                        vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                        vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                        " - UserIndex: " & Userindex & " - producido al manejar el paquete: " & CStr(packetID))
        Call CloseSocket(Userindex)
        HandleIncomingData = True
    Else
        'Flush buffer - send everything that has been written
        Call FlushBuffer(Userindex)
        HandleIncomingData = True
    End If
End Function

''
' Handles the "LoginExistingChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If SeguridadAlkon Then
    If UserList(Userindex).incomingData.Length < 68 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#Else
    If UserList(Userindex).incomingData.Length < 22 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String
    Dim UserAccount As String
    Dim password As String
    Dim version As String
    UserAccount = buffer.ReadASCIIString()
    UserName = buffer.ReadASCIIString()
    
    UserList(Userindex).iServer = 0
    UserList(Userindex).iCliente = 0
    UserList(Userindex).clave = StrConv("damn" & StrReverse(UCase$(UserName)) & "you", vbFromUnicode)

    password = buffer.ReadASCIIStringFixed(32)

    
    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    If Not AsciiValidos(UserAccount) Then
        Call WriteErrorMsg(Userindex, "Cuenta inválida.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        
        Exit Sub
    End If
    
    If Not AsciiValidos(UserName) Then
        Call WriteErrorMsg(Userindex, "Nombre inválido.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        
        Exit Sub
    End If
    If Not CuentaExiste(UserAccount) Then
        Call WriteErrorMsg(Userindex, "La cuenta no existe.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        
        Exit Sub
    End If
    If Not PersonajeExiste(UserName) Then
        Call WriteErrorMsg(Userindex, "El personaje no existe.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        
        Exit Sub
    End If
    

#If SeguridadAlkon Then
    If Not MD5ok(buffer.ReadASCIIStringFixed(16)) Then
        Call WriteErrorMsg(Userindex, "El cliente está dañado, por favor descarguelo nuevamente desde www.nuevoargentum.com")
    Else
#End If
        
        If BANCheck(UserName) Then
            Dim Dias As Long
            Dias = GetByCampo("SELECT DATEDIFF(BanTime, NOW()) as Dias FROM pjs WHERE Nombre=" & Comillas(UserName), "Dias")
            If Dias < -4000 Then
                Call WriteErrorMsg(Userindex, "Se te ha prohibido la entrada de manera indefinida a AoYind debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.nuevoargentum.com")
            ElseIf Dias > 0 Then
                Call WriteErrorMsg(Userindex, "Se te ha prohibido la entrada a AoYind por " & Dias & " debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.nuevoargentum.com")
            Else
                Call modMySQL.Execute("UPDATE pjs SET BanTime=20000101, Ban=0 WHERE Nombre=" & Comillas(UserName))
            End If
        ElseIf Not VersionOK(version) Then
            Call WriteErrorMsg(Userindex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en www.nuevoargentum.com")
        Else
            Call ConnectUser(Userindex, UserAccount, UserName, password)
        End If
#If SeguridadAlkon Then
    End If
#End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleOpenAccount(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If SeguridadAlkon Then
    If UserList(Userindex).incomingData.Length < 68 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#Else
    If UserList(Userindex).incomingData.Length < 37 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String
    Dim password As String
    Dim version As String
    

    UserName = buffer.ReadASCIIString()
    

    password = buffer.ReadASCIIStringFixed(32)

    
    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    If Not AsciiValidos(UserName) Then
        Call WriteErrorMsg(Userindex, "Nombre inválido.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        
        Exit Sub
    End If
    
    If Not CuentaExiste(UserName) Then
        Call WriteErrorMsg(Userindex, "La cuenta no existe.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        
        Exit Sub
    End If
    
        
    If BANCuentaCheck(UserName) Then
        Call WriteErrorMsg(Userindex, "Se te ha prohibido la entrada a Argentum debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.nuevoargentum.com.ar")
    ElseIf Not VersionOK(version) Then
        Call WriteErrorMsg(Userindex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en www.nuevoargentum.com.ar")
    Else
        Call WriteOpenAccount(Userindex, UserName, password)
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


Private Sub HandleBorrarPJ(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If SeguridadAlkon Then
    If UserList(Userindex).incomingData.Length < 68 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#Else
    If UserList(Userindex).incomingData.Length < 37 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String
    Dim UserPJ As String
    Dim password As String
    Dim version As String
    

    UserName = buffer.ReadASCIIString()
    UserPJ = buffer.ReadASCIIString()
    password = buffer.ReadASCIIStringFixed(32)
    
    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    If Not AsciiValidos(UserName) Then
        Call WriteErrorMsg(Userindex, "Nombre inválido.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        
        Exit Sub
    End If
    
    If Not CuentaExiste(UserName) Then
        Call WriteErrorMsg(Userindex, "La cuenta no existe.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        
        Exit Sub
    End If
        
    If BANCuentaCheck(UserName) Then
        Call WriteErrorMsg(Userindex, "Esta cuenta se encuentra baneada")
        Call FlushBuffer(Userindex)
    ElseIf Not VersionOK(version) Then
        Call WriteErrorMsg(Userindex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en www.nuevoargentum.com")
        Call FlushBuffer(Userindex)
    Else
        Call WriteBorrarPJ(Userindex, UserName, UserPJ, password)
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub



''
' Handles the "LoginNewChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If SeguridadAlkon Then
    If UserList(Userindex).incomingData.Length < 81 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#Else
    If UserList(Userindex).incomingData.Length < 49 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String
    Dim UserAccount As String
    Dim password As String
    Dim version As String
    Dim race As eRaza
    Dim gender As eGenero
    Dim homeland As eCiudad
    Dim Class As eClass
    
#If SeguridadAlkon Then
    Dim MD5 As String
#End If
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(Userindex, "La creacion de personajes en este servidor se ha deshabilitado.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        
        Exit Sub
    End If
    
    If ServerSoloGMs <> 0 Then
        Call WriteErrorMsg(Userindex, "Servidor restringido a administradores. Consulte la página oficial o el foro oficial para mas información.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        
        Exit Sub
    End If
    
    If aClon.MaxPersonajes(UserList(Userindex).ip) Then
        Call WriteErrorMsg(Userindex, "Has creado demasiados personajes.")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        
        Exit Sub
    End If
    UserAccount = buffer.ReadASCIIString()
    UserName = buffer.ReadASCIIString()
    
    password = buffer.ReadASCIIStringFixed(32)

    Dim idCuenta As Long
    
    idCuenta = CLngNull(GetByCampo("SELECT Id FROM cuentas WHERE Nombre=" & Comillas(UserAccount) & " AND Password=" & Comillas(password), "Id"))
    
    If idCuenta = 0 Then
        Call WriteErrorMsg(Userindex, "Los datos de la cuenta son incorrectos")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
        
        Exit Sub
    End If
    
    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
#If SeguridadAlkon Then
    MD5 = buffer.ReadASCIIStringFixed(16)
#End If
    
    race = buffer.ReadByte()
    gender = buffer.ReadByte()
    Class = buffer.ReadByte()
    homeland = buffer.ReadByte()
    
#If SeguridadAlkon Then
    If Not MD5ok(MD5) Then
        Call WriteErrorMsg(Userindex, "El cliente está dañado, por favor descarguelo nuevamente desde www.argentumonline.com.ar")
    Else
#End If
        
        If Not VersionOK(version) Then
            Call WriteErrorMsg(Userindex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en www.argentumonline.com.ar")
        Else
            Call ConnectNewUser(Userindex, idCuenta, UserAccount, UserName, password, race, gender, Class, homeland)
        End If
#If SeguridadAlkon Then
    End If
#End If

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleLoginNewAccount(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
#If SeguridadAlkon Then
    If UserList(Userindex).incomingData.Length < 81 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#Else
    If UserList(Userindex).incomingData.Length < 49 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
#End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
  
    Dim UserAccount As String
    Dim UserEmail As String
    Dim password As String
    Dim version As String
       
    #If SeguridadAlkon Then
        Dim MD5 As String
    #End If
       
    
    UserAccount = buffer.ReadASCIIString()
    UserEmail = buffer.ReadASCIIString()
    password = buffer.ReadASCIIStringFixed(32)

    'Convert version number to string
    version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    #If SeguridadAlkon Then
        MD5 = buffer.ReadASCIIStringFixed(16)
    #End If
    
    #If SeguridadAlkon Then
        If Not MD5ok(MD5) Then
            Call WriteErrorMsg(Userindex, "El cliente está dañado, por favor descarguelo nuevamente desde www.argentumonline.com.ar")
        Else
    #End If
        
        If Not VersionOK(version) Then
            Call WriteErrorMsg(Userindex, "Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en www.argentumonline.com.ar")
        Else
            Call CreateNewAccount(Userindex, UserAccount, UserEmail, password)
        End If
    
    #If SeguridadAlkon Then
        End If
    #End If

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Dijo: " & Chat)
        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(Userindex)
                    Call WriteConsoleMsg(Userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(Userindex, UserList(Userindex).Char.CharIndex, False)
                    Call WriteConsoleMsg(Userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .flags.Muerto = 1 Then
                Call SendData(SendTarget.ToDeadArea, Userindex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
            Else
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, .flags.ChatColor))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If UserList(Userindex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos.", FontTypeNames.FONTTYPE_INFO)
        Else
            '[Consejeros & GMs]
            If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                Call LogGM(.Name, "Grito: " & Chat)
            End If
            
            'I see you....
            If .flags.Oculto > 0 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                If .flags.Navegando = 1 Then
                    If .clase = eClass.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
                        Call ToggleBoatBody(Userindex)
                        Call WriteConsoleMsg(Userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.heading, NingunArma, _
                                            NingunEscudo, NingunCasco)
                    End If
                Else
                    If .flags.invisible = 0 Then
                        Call UsUaRiOs.SetInvisible(Userindex, .Char.CharIndex, False)
                        Call WriteConsoleMsg(Userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
            
            If LenB(Chat) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(Chat)
                
                If .flags.Privilegios And PlayerType.User Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, vbRed))
                Else
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_GM_YELL))
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 28/05/2009
'28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        Dim targetCharIndex As Integer
        Dim TargetUserIndex As Integer
        Dim targetPriv As PlayerType
        
        targetCharIndex = buffer.ReadInteger()
        Chat = buffer.ReadASCIIString()
        
        TargetUserIndex = CharIndexToUserIndex(targetCharIndex)
        
        If .flags.Muerto Then
            Call WriteConsoleMsg(Userindex, "¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
        Else
            If TargetUserIndex = INVALID_INDEX Then
                Call WriteConsoleMsg(Userindex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
            Else
                targetPriv = UserList(TargetUserIndex).flags.Privilegios
                'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
                If (targetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Then
                    ' Controlamos que no este invisible
                    If UserList(TargetUserIndex).flags.AdminInvisible <> 1 Then
                        Call WriteConsoleMsg(Userindex, "No puedes susurrarle a los Dioses y Admins.", FontTypeNames.FONTTYPE_INFO)
                    End If
                'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
                ElseIf (.flags.Privilegios And PlayerType.User) <> 0 And (Not targetPriv And PlayerType.User) <> 0 Then
                    ' Controlamos que no este invisible
                    If UserList(TargetUserIndex).flags.AdminInvisible <> 1 Then
                        Call WriteConsoleMsg(Userindex, "No puedes susurrarle a los GMs.", FontTypeNames.FONTTYPE_INFO)
                    End If
                ElseIf Not EstaPCarea(Userindex, TargetUserIndex) Then
                    Call WriteConsoleMsg(Userindex, "Estas muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                
                Else
                    '[Consejeros & GMs]
                    If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                        Call LogGM(.Name, "Le dijo a '" & UserList(TargetUserIndex).Name & "' " & Chat)
                    End If
                    
                    If LenB(Chat) <> 0 Then
                        'Analize chat...
                        Call Statistics.ParseChat(Chat)
                        
                        Call WriteChatOverHead(Userindex, Chat, .Char.CharIndex, vbBlue)
                        Call WriteChatOverHead(TargetUserIndex, Chat, .Char.CharIndex, vbBlue)
                        Call FlushBuffer(TargetUserIndex)
                        
                        '[CDT 17-02-2004]
                        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                            Call SendData(SendTarget.ToAdminsAreaButConsejeros, Userindex, PrepareMessageChatOverHead("a " & UserList(TargetUserIndex).Name & "> " & Chat, .Char.CharIndex, vbYellow))
                        End If
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'If UserList(UserIndex).flags.Wait = 1 Then Exit Sub
    
    Dim dummy As Long
    Dim TempTick As Long
    Dim heading As eHeading
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        heading = .incomingData.ReadByte()
        
        .flags.Movimiento = .flags.Movimiento + IIf(heading <> .Char.heading, 4, 2)
        If .flags.Movimiento < 7 Then .flags.Movimiento = 7
        If .flags.Movimiento > 24 Then .flags.Movimiento = 24
        
        'Prevent SpeedHack
        If .flags.TimesWalk >= 24 Then
            TempTick = (GetTickCount() And &H7FFFFFFF)
            dummy = (TempTick - .flags.StartWalk)
            Debug.Print dummy
            
            ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
            '(it's about 193 ms per step against the over 200 needed in perfect conditions)
            If dummy < 3850 Then
                If TempTick - .flags.CountSH > 30000 Then
                    .flags.CountSH = 0
                End If
                
                If Not .flags.CountSH = 0 Then
                    If dummy <> 0 Then _
                        dummy = 126000 \ dummy
                    
                    Call LogHackAttemp("Tramposo SH: " & .Name & " , " & dummy)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha sido echado por el servidor por posible uso de Speed Hack.", FontTypeNames.FONTTYPE_SERVER))
                    Call CloseSocket(Userindex)
                    
                    Exit Sub
                Else
                    .flags.CountSH = TempTick
                End If
            End If
            .flags.StartWalk = TempTick
            .flags.TimesWalk = 0
        End If
        
        .flags.TimesWalk = .flags.TimesWalk + 1
        
        'If exiting, cancel
        Call CancelExit(Userindex)
        
        'TODO: Debería decirle por consola que no puede?
        'Esta usando el /HOGAR, no se puede mover
        If .flags.Traveling = 1 Then
            .flags.Traveling = 0
            .Counters.goHome = 0
            Call WriteGotHome(Userindex, False)
        End If
        
        
        
        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
                .Char.Loops = 0
                
                Call WriteMeditateToggle(Userindex)
                Call WriteConsoleMsg(Userindex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Else
                'Move user
                Call MoveUserChar(Userindex, heading)
                
                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                    
                    Call WriteRestOK(Userindex)
                    Call WriteConsoleMsg(Userindex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        Else    'paralized
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteConsoleMsg(Userindex, "No podes moverte porque estas paralizado.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.CountSH = 0
        End If
        
        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
            If .clase <> eClass.Thief And .clase <> eClass.Bandit Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
            
                If .flags.Navegando = 1 Then
                    If .clase = eClass.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
                        Call ToggleBoatBody(Userindex)
                        Call WriteConsoleMsg(Userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                    End If
                Else
                    'If not under a spell effect, show char
                    If .flags.invisible = 0 Then
                        Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                        Call UsUaRiOs.SetInvisible(Userindex, .Char.CharIndex, False)
                    End If
                End If
            End If
        End If
        .Counters.PiqueteC = 0
        Call CheckZona(Userindex)
        Call DoTileEvents(Userindex, .Pos.map, .Pos.X, .Pos.Y)
    End With
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    UserList(Userindex).incomingData.ReadByte
    
    Call WritePosUpdate(Userindex)
End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/01/08
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
    
        Dim heading As Byte
        'Remove packet ID
        Call .incomingData.ReadByte
        heading = .incomingData.ReadByte
        
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡No podes atacar a nadie porque estas muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(Userindex, eMessages.mUserMuerto)
            Exit Sub
        End If
               
        'If user meditates, can't attack
        If .flags.Meditando Then
            Exit Sub
        End If
        
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(Userindex, "No podés usar así esta arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        If heading <> .Char.heading Then
            .Char.heading = heading
            Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
        'If exiting, cancel
        Call CancelExit(Userindex)
        
        'Attack!
        Call UsuarioAtaca(Userindex)
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(Userindex)
                    Call WriteConsoleMsg(Userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(Userindex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(Userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With
End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then Exit Sub
        
        'Lower rank administrators can't pick up items
        If .flags.Privilegios And PlayerType.Consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteConsoleMsg(Userindex, "No puedes tomar ningun objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        Call GetObj(Userindex)
    End With
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Seguro Then
            Call WriteSafeModeOff(Userindex)
        Else
            Call WriteSafeModeOn(Userindex)
        End If
        
        .flags.Seguro = Not .flags.Seguro
    End With
End Sub

''
' Handles the "ResuscitationSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResuscitationToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Rapsodius
'Creation Date: 10/10/07
'***************************************************
    With UserList(Userindex)
        Call .incomingData.ReadByte
        
        .flags.SeguroResu = Not .flags.SeguroResu
        
        If .flags.SeguroResu Then
            Call WriteResuscitationSafeOn(Userindex)
        Else
            Call WriteResuscitationSafeOff(Userindex)
        End If
    End With
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    UserList(Userindex).incomingData.ReadByte
    
    Call modGuilds.SendGuildLeaderInfo(Userindex)
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    Call WriteAttributes(Userindex, False)
End Sub

''
' Handles the "RequestFame" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestFame(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    Call EnviarFama(Userindex)
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    Call WriteSendSkills(Userindex)
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    Call WriteMiniStats(Userindex)
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    'User quits commerce mode
    UserList(Userindex).flags.Comerciando = False
    Call WriteCommerceEnd(Userindex)
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 And UserList(.ComUsu.DestUsu).ComUsu.DestUsu = Userindex Then
            Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(.ComUsu.DestUsu)
            
            'Send data in the outgoing buffer of the other user
            Call FlushBuffer(.ComUsu.DestUsu)
        End If
        
        Call FinComerciarUsu(Userindex)
    End With
End Sub

''
' Handles the "UserCommerceConfirm" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUserCommerceConfirm(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte

    'Validate the commerce
    If PuedeSeguirComerciando(Userindex) Then
        'Tell the other user the confirmation of the offer
        Call WriteUserOfferConfirm(UserList(Userindex).ComUsu.DestUsu)
        UserList(Userindex).ComUsu.Confirmo = True
    End If
    
End Sub
Public Sub WriteUserOfferConfirm(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UserOfferConfirm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserOfferConfirm)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Private Sub HandleCommerceChat(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            If PuedeSeguirComerciando(Userindex) Then
                'Analize chat...
                Call Statistics.ParseChat(Chat)
                
                Chat = UserList(Userindex).Name & "> " & Chat
                Call WriteCommerceChat(Userindex, Chat, FontTypeNames.FONTTYPE_PARTY)
                Call WriteCommerceChat(UserList(Userindex).ComUsu.DestUsu, Chat, FontTypeNames.FONTTYPE_PARTY)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(Userindex)
    End With
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    'Trade accepted
    Call AceptarComercioUsu(Userindex)
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim otherUser As Integer
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        otherUser = .ComUsu.DestUsu
        
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)
                
                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(otherUser)
            End If
        End If
        
        Call WriteConsoleMsg(Userindex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(Userindex)
    End With
End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 8 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim Slot As Byte
    Dim Amount As Integer
    Dim DropX As Integer
    Dim DropY As Integer
    Dim Trigger As Byte
    Dim tmpInt As Integer
    Dim ObjIndex As Integer
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        DropX = .incomingData.ReadInteger()
        DropY = .incomingData.ReadInteger()
        
        If .Counters.Pena > 0 Then
            Call WriteConsoleMsg(Userindex, "No se puede tirar objetos en prisión.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.Embarcado Then
            Call WriteConsoleMsg(Userindex, "No se puede tirar objetos desde el barco.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Navegando = 1 Or _
           .flags.Muerto = 1 Or _
           ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        If Amount < 0 Then Amount = 0
        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If Amount > 10000 Then Exit Sub 'Don't drop too much gold

            Call TirarOro(Amount, Userindex)
            
            Call WriteUpdateGold(Userindex)
        Else
            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then
                    Exit Sub
                End If
                If DropX = 0 Then
                    DropX = .Pos.X
                    DropY = .Pos.Y
                Else
                    tmpInt = MapData(.Pos.map).Tile(DropX, DropY).NpcIndex
                    If tmpInt > 0 Then
                        If Npclist(tmpInt).NpcType = Banquero Then
                        
                            If Abs(.Pos.X - DropX) > 3 Or Abs(.Pos.Y - DropY) > 3 Then
                                Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                            ElseIf .SalaIndex > 0 Then
                                Call WriteConsoleMsg(Userindex, "No se puede depositar objetos durante un reto.", FontTypeNames.FONTTYPE_INFO)
                            Else
                        
                                If Amount > UserList(Userindex).Invent.Object(Slot).Amount Then Amount = UserList(Userindex).Invent.Object(Slot).Amount
                                ObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
                                'Agregamos el obj que deposita al banco
                                If UserDejaObj(Userindex, CInt(Slot), Amount) Then
                                    'Actualizamos el inventario del usuario
                                    Call UpdateUserInv(False, Userindex, Slot)
                                    Call WriteChatOverHead(Userindex, "Has depositado " & Amount & " " & ObjData(ObjIndex).Name & ".", Npclist(tmpInt).Char.CharIndex, vbWhite)

                                End If
                            
                            End If
                            Exit Sub
                            
                        End If
                    End If
                    If Zonas(UserList(Userindex).zona).Segura Then
                        Call WriteConsoleMsg(Userindex, "No se puede arrojar objetos en zonas seguras.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    If DropX > .Pos.X + 3 Then DropX = .Pos.X + 3
                    If DropX < .Pos.X - 3 Then DropX = .Pos.X - 3
                    If DropY > .Pos.Y + 3 Then DropY = .Pos.Y + 3
                    If DropY < .Pos.Y - 3 Then DropY = .Pos.Y - 3
                    
                    Trigger = MapData(.Pos.map).Tile(DropX, DropY).Trigger
                    tmpInt = MapData(.Pos.map).Tile(DropX, DropY).NpcIndex
                    If MapData(.Pos.map).Tile(DropX, DropY).Blocked Then
                        Call WriteConsoleMsg(Userindex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    ElseIf tmpInt > 0 Then
                        Call WriteConsoleMsg(Userindex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    ElseIf Trigger = eTrigger.trigger_2 Or Trigger = eTrigger.BAJOTECHO Or Trigger = eTrigger.AUTORESU Or Trigger = eTrigger.ZONASEGURA Then
                        Call WriteConsoleMsg(Userindex, "No se puede arrojar objetos dentro de las casas.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                Call DropObj(Userindex, Slot, Amount, .Pos.map, DropX, DropY)
            End If
        End If
    End With
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Spell As Byte
        
        Spell = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estas muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        

        If Spell < 1 Then
            .flags.hechizo = 0
            Exit Sub
        ElseIf Spell > MAXUSERHECHIZOS Then
            .flags.hechizo = 0
            Exit Sub
        End If
        
        .flags.hechizo = .Stats.UserHechizos(Spell)
    End With
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Integer
        Dim Y As Integer
        
        X = .ReadInteger()
        Y = .ReadInteger()
        
        Call LookatTile(Userindex, UserList(Userindex).Pos.map, X, Y)
    End With
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Integer
        Dim Y As Integer
        
        X = .ReadInteger()
        Y = .ReadInteger()
        
        Call Accion(Userindex, UserList(Userindex).Pos.map, X, Y)
    End With
End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Skill As eSkill
        
        Skill = .incomingData.ReadByte()
        
        If UserList(Userindex).flags.Muerto = 1 Then Exit Sub
        
        'If exiting, cancel
        Call CancelExit(Userindex)
        
        Select Case Skill
            Case Robar, Magia, Domar
                Call WriteWorkRequestTarget(Userindex, Skill)
            Case Ocultarse
                If .flags.Navegando = 1 Then
                    If .clase <> eClass.Pirat Then
                        '[CDT 17-02-2004]
                        If Not .flags.UltimoMensaje = 3 Then
                            Call WriteConsoleMsg(Userindex, "No puedes ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
                            .flags.UltimoMensaje = 3
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                End If
                
                If .flags.Oculto = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(Userindex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                If Zonas(.zona).Segura = 1 Or MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger = eTrigger.ZONASEGURA Then
                    If Not .flags.UltimoMensaje = 11 Then
                        Call WriteConsoleMsg(Userindex, "¡No está permitido ocultarse en zonas seguras!.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 11
                    End If
                    Exit Sub
                End If
                'No usar invi mapas InviSinEfecto
                If Zonas(.zona).InviSinEfecto > 0 Then
                    If Not .flags.UltimoMensaje = 11 Then
                        Call WriteConsoleMsg(Userindex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 11
                    End If
                    Exit Sub
                End If
                
                Call DoOcultarse(Userindex)
        End Select
    End With
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call SendData(SendTarget.ToAdmins, Userindex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos", FontTypeNames.FONTTYPE_VENENO))
        Call WriteErrorMsg(Userindex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros")
        Call FlushBuffer(Userindex)
        Call CloseSocket(Userindex)
    End With
End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        
        Slot = .incomingData.ReadByte()
        
        If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        End If
        
        If .flags.Meditando Then
            Exit Sub    'The error message should have been provided by the client.
        End If
        
        Call UseInvItem(Userindex, Slot)
    End With
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        
        Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
        Call HerreroConstruirItem(Userindex, Item)
    End With
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        
        Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
        Call CarpinteroConstruirItem(Userindex, Item)
    End With
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Integer
        Dim Y As Integer
        Dim Skill As eSkill
        Dim DummyInt As Integer
        Dim tU As Integer   'Target user
        Dim tN As Integer   'Target NPC
        
        X = .incomingData.ReadInteger()
        Y = .incomingData.ReadInteger()
        
        Skill = .incomingData.ReadByte()
        
        
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando _
        Or Not InMapBounds(.Pos.map, X, Y) Then Exit Sub

        If Not InRangoVision(Userindex, X, Y) Then
            Call WritePosUpdate(Userindex)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(Userindex)
        
        Select Case Skill
            Case eSkill.Proyectiles
'
'                'Check attack interval
'                If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub
'                'Check Magic interval
'                If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Sub
'                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(Userindex) Then Exit Sub
                
                
                Call LanzarProyectil(Userindex, X, Y)
                
            
            Case eSkill.Magia
                'Check the map allows spells to be casted.
                If Zonas(.zona).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(Userindex, "Una fuerza oscura te impide canalizar tu energía", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                'Target whatever is in that tile
                Call LookatTile(Userindex, .Pos.map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .ip & " a la posicion (" & .Pos.map & "/" & X & "/" & Y & ")")
                    Exit Sub
                End If
                
                'Check bow's interval
                'If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                
                'Check Spell-Hit interval
'                If Not IntervaloPermiteGolpeMagia(UserIndex) Then
'                    'Check Magic interval
                    If Not IntervaloPermiteLanzarSpell(Userindex) Then Exit Sub
'                    End If
'                End If
                
                
                'Check intervals and cast
                If .flags.hechizo > 0 Then
                    Call LanzarHechizo(.flags.hechizo, Userindex)
                    .flags.hechizo = 0
                Else
                    Call WriteConsoleMsg(Userindex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Pesca
                DummyInt = .Invent.WeaponEqpObjIndex
                If DummyInt = 0 Then Exit Sub
                
                'Check interval
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
                
                If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger = 1 Then
                    Call WriteConsoleMsg(Userindex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .flags.Embarcado = 1 Then
                    Call WriteConsoleMsg(Userindex, "No puedes pescar arriba del barco viajero.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .flags.Nadando = True Then
                    Call WriteConsoleMsg(Userindex, "No puedes pescar nadando...", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .flags.Equitando = True Then
                    Call WriteConsoleMsg(Userindex, "No puedes pescar montado...", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If HayAgua(.Pos.map, X, Y) Then
                    Select Case DummyInt
                        Case CAÑA_PESCA
                            Call DoPescar(Userindex)
                        
                        Case RED_PESCA
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(Userindex, "Estás demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            Call DoPescarRed(Userindex)
                        
                        Case Else
                            Exit Sub    'Invalid item!
                    End Select
                    
                    'Play sound!
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                Else
                    Call WriteConsoleMsg(Userindex, "No hay agua donde pescar. Busca un lago, río o mar.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Robar
                'Does the map allow us to steal here?
                If Zonas(.zona).Segura = 0 Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(Userindex, UserList(Userindex).Pos.map, X, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> Userindex Then
                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                 If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                     Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                     Exit Sub
                                 End If
                                 
                                 '17/09/02
                                 'Check the trigger
                                 If MapData(UserList(tU).Pos.map).Tile(X, Y).Trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(Userindex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(Userindex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 Call DoRobar(Userindex, tU)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(Userindex, "No a quien robarle!.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "¡No puedes robar en zonas seguras!.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Talar
                'Check interval
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
                
                If .Invent.WeaponEqpObjIndex = 0 Then
                    Call WriteConsoleMsg(Userindex, "Deberías equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .Invent.WeaponEqpObjIndex <> HACHA_LEÑADOR And _
                    .Invent.WeaponEqpObjIndex <> HACHA_LEÑA_ELFICA Then
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                DummyInt = MapData(.Pos.map).Tile(X, Y).ObjInfo.ObjIndex
                
                If DummyInt > 0 Then
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If .Pos.X = X And .Pos.Y = Y Then
                        Call WriteConsoleMsg(Userindex, "No podés talar desde allí.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If .flags.Equitando = True Then
                        Call WriteConsoleMsg(Userindex, "No puedes talar montado...", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(DummyInt).OBJType = eOBJType.otArboles And (.Invent.WeaponEqpObjIndex = HACHA_LEÑADOR Or .Invent.WeaponEqpObjIndex = HACHA_LEÑA_ELFICA) Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                        Call DoTalar(Userindex)
                    ElseIf ObjData(DummyInt).OBJType = eOBJType.otArbolElfico And .Invent.WeaponEqpObjIndex = HACHA_LEÑA_ELFICA Then
                        If .Invent.WeaponEqpObjIndex = HACHA_LEÑA_ELFICA Then
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                            Call DoTalar(Userindex, True)
                        Else
                            Call WriteConsoleMsg(Userindex, "El hacha utilizado no es suficientemente poderosa.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Mineria
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
                                
                If .Invent.WeaponEqpObjIndex = 0 Then Exit Sub
                
                If .Invent.WeaponEqpObjIndex <> PIQUETE_MINERO Then
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                'Target whatever is in the tile
                Call LookatTile(Userindex, .Pos.map, X, Y)
                
                DummyInt = MapData(.Pos.map).Tile(X, Y).ObjInfo.ObjIndex
                
                If DummyInt > 0 Then
                    'Check distance
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If .flags.Equitando = True Then
                        Call WriteConsoleMsg(Userindex, "No puedes minar montado...", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    DummyInt = MapData(.Pos.map).Tile(X, Y).ObjInfo.ObjIndex 'CHECK
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
                        Call DoMineria(Userindex)
                    Else
                        Call WriteConsoleMsg(Userindex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                
                'Target whatever is that tile
                Call LookatTile(Userindex, .Pos.map, X, Y)
                tN = .flags.TargetNPC
                
                If tN > 0 Then
                    If Npclist(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(Userindex, "No podés domar una criatura que está luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoDomar(Userindex, tN)
                    Else
                        Call WriteConsoleMsg(Userindex, "No podés domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No hay ninguna criatura allí!.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                'Check interval
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
                
                'Check there is a proper item there
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
                        'Validate other items
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > MAX_INVENTORY_SLOTS Then
                            Exit Sub
                        End If
                        
                        ''chequeamos que no se zarpe duplicando oro
                        If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                            If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                                Call WriteConsoleMsg(Userindex, "No tienes más minerales", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            ''FUISTE
                            Call WriteErrorMsg(Userindex, "Has sido expulsado por el sistema anti cheats.")
                            Call FlushBuffer(Userindex)
                            Call CloseSocket(Userindex)
                            Exit Sub
                        End If
                        
                        Call FundirMineral(Userindex)
                    Else
                        Call WriteConsoleMsg(Userindex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Herreria
                'Target wehatever is in that tile
                Call LookatTile(Userindex, .Pos.map, X, Y)
                
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(Userindex)
                        Call EnivarArmadurasConstruibles(Userindex)
                        Call WriteShowBlacksmithForm(Userindex)
                    Else
                        Call WriteConsoleMsg(Userindex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "Ahí no hay ningún yunque.", FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/11/09
'05/11/09: Pato - Ahora se quitan los espacios del principio y del fin del nombre del clan
'***************************************************
    If UserList(Userindex).incomingData.Length < 9 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim desc As String
        Dim GuildName As String
        Dim site As String
        Dim Codex() As String
        Dim errorStr As String
        
        desc = buffer.ReadASCIIString()
        GuildName = Trim$(buffer.ReadASCIIString())
        site = buffer.ReadASCIIString()
        Codex = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        If modGuilds.CrearNuevoClan(Userindex, desc, GuildName, site, Codex, .FundandoGuildAlineacion, errorStr) Then
            Call SendData(SendTarget.ToAll, Userindex, PrepareMessageConsoleMsg(.Name & " fundó el clan " & GuildName & " de alineación " & modGuilds.GuildAlignment(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))

            
            'Update tag
             Call RefreshCharStatus(Userindex)
        Else
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim spellSlot As Byte
        Dim Spell As Integer
        
        spellSlot = .incomingData.ReadByte()
        
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(Userindex, "¡Primero selecciona el hechizo.!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(Userindex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                               & "Nombre:" & .Nombre & vbCrLf _
                                               & "Descripción:" & .desc & vbCrLf _
                                               & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                               & "Mana necesario: " & .ManaRequerido & vbCrLf _
                                               & "Stamina necesaria: " & .StaRequerido & vbCrLf _
                                               & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
            End With
        End If
    End With
End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim itemSlot As Byte
        
        itemSlot = .incomingData.ReadByte()
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then Exit Sub
        
        'Validate item slot
        If itemSlot > MAX_INVENTORY_SLOTS Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
        Call EquiparInvItem(Userindex, itemSlot)
    End With
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/28/2008
'Last Modified By: NicoNZ
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
' 06/28/2008: NicoNZ - Sólo se puede cambiar si está inmovilizado.
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim heading As eHeading
        Dim posX As Integer
        Dim posY As Integer
                
        heading = .incomingData.ReadByte()
        
        If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then
            Select Case heading
                Case eHeading.NORTH
                    posY = -1
                Case eHeading.EAST
                    posX = 1
                Case eHeading.SOUTH
                    posY = 1
                Case eHeading.WEST
                    posX = -1
            End Select
            
                If LegalPos(.Pos.map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                    Exit Sub
                End If
        End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If heading > 0 And heading < 5 And heading <> .Char.heading Then
            .Char.heading = heading
            Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 1 + NUMSKILLS Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim i As Long
        Dim Count As Integer
        Dim points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            points(i) = .incomingData.ReadByte()
            
            If points(i) < 0 Then
                Call LogHackAttemp(.Name & " IP:" & .ip & " trató de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(Userindex)
                Exit Sub
            End If
            
            Count = Count + points(i)
        Next i
        
        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.Name & " IP:" & .ip & " trató de hackear los skills.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        With .Stats
            For i = 1 To NUMSKILLS
                .SkillPts = .SkillPts - points(i)
                .UserSkills(i) = .UserSkills(i) + points(i)
                
                'Client should prevent this, but just in case...
                If .UserSkills(i) > 100 Then
                    .SkillPts = .SkillPts + .UserSkills(i) - 100
                    .UserSkills(i) = 100
                End If
            Next i
        End With
    End With
End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim SpawnedNpc As Integer
        Dim petIndex As Byte
        
        petIndex = .incomingData.ReadByte()
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNPC).NpcType <> eNPCType.Entrenador Then Exit Sub
        
        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If petIndex > 0 And petIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(petIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False, UserList(Userindex).zona)
                
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1
                End If
            End If
        Else
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
        End If
    End With
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
            
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(Userindex, "No estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'User compra el item
        Call Comercio(eModoComercio.Compra, Userindex, .flags.TargetNPC, Slot, Amount)
    End With
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿Es el banquero?
        If Npclist(.flags.TargetNPC).NpcType <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User retira el item del slot
        Call UserRetiraItem(Userindex, Slot, Amount)
    End With
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, Userindex, .flags.TargetNPC, Slot, Amount)
    End With
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .SalaIndex > 0 Then
            Call WriteConsoleMsg(Userindex, "No se puede depositar objetos durante un reto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '¿El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).NpcType <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User deposita el item del slot rdata
        Call UserDepositaItem(Userindex, Slot, Amount)
    End With
End Sub


''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Dir As Integer
        
        If .ReadBoolean() Then
            Dir = 1
        Else
            Dir = -1
        End If
        
        Call DesplazarHechizo(Userindex, Dir, .ReadByte())
    End With
End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal Userindex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/14/09
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Dir As Integer
        Dim Slot As Byte
        Dim TempItem As Obj
        
        If .ReadBoolean() Then
            Dir = 1
        Else
            Dir = -1
        End If
        
        Slot = .ReadByte()
    End With
    
    
    
    With UserList(Userindex)
        'TempItem.ObjIndex = .BancoInvent.Object(Slot).ObjIndex
        'TempItem.Amount = .BancoInvent.Object(Slot).Amount
        
        If Dir = 1 Then 'Mover arriba
        
            Call IntercambiarBanco(Userindex, Slot, Slot - 1)
        
            '.BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
            '.BancoInvent.Object(Slot - 1).ObjIndex = TempItem.ObjIndex
            '.BancoInvent.Object(Slot - 1).Amount = TempItem.Amount
        Else 'mover abajo
        
            Call IntercambiarBanco(Userindex, Slot, Slot + 1)
        
            '.BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
            '.BancoInvent.Object(Slot + 1).ObjIndex = TempItem.ObjIndex
            '.BancoInvent.Object(Slot + 1).Amount = TempItem.Amount
        End If
    End With
    
    Call UpdateBanUserInv(True, Userindex, 0)
    Call UpdateVentanaBanco(Userindex)

End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim desc As String
        Dim Codex() As String
        
        desc = buffer.ReadASCIIString()
        Codex = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        Call modGuilds.ChangeCodexAndDesc(desc, Codex, .GuildIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 7 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        Dim Slot As Byte
        Dim tUser As Integer
        Dim OfferSlot As Byte
        Dim ObjIndex As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadLong()
        OfferSlot = .incomingData.ReadByte()
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        ' If he's already confirmed his offer, but now tries to change it, then he's cheating
        If UserList(Userindex).ComUsu.Confirmo = True Then
            
            ' Finish the trade
            Call FinComerciarUsu(Userindex)
        
            If tUser <= 0 Or tUser > MaxUsers Then
                Call FinComerciarUsu(tUser)
                Call Protocol.FlushBuffer(tUser)
            End If
        
            Exit Sub
        End If
        
        'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
        If ((Slot < 0 Or Slot > UserList(Userindex).CurrentInventorySlots) And Slot <> FLAGORO) Then Exit Sub
        
        'If OfferSlot is invalid, then ignore it.
        If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 1 Then Exit Sub
        
        ' Can be negative if substracted from the offer, but never 0.
        If Amount = 0 Then Exit Sub
        
        'Has he got enough??
        If Slot = FLAGORO Then
            ' Can't offer more than he has
            If Amount > .Stats.GLD - .ComUsu.GoldAmount Then
                Call WriteCommerceChat(Userindex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
        Else
            'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
            If Slot <> 0 Then ObjIndex = .Invent.Object(Slot).ObjIndex
            ' Can't offer more than he has
            If Not TieneObjetos(ObjIndex, _
                TotalOfferItems(ObjIndex, Userindex) + Amount, Userindex) Then
                
                Call WriteCommerceChat(Userindex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            
            If ItemNewbie(ObjIndex) Then
                Call WriteCancelOfferItem(Userindex, OfferSlot)
                Exit Sub
            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteCommerceChat(Userindex, "No puedes vender tu barco mientras lo estés usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            
        End If
        
        
                
        Call AgregarOferta(Userindex, OfferSlot, ObjIndex, Amount, Slot = FLAGORO)
        
        Call EnviarOferta(tUser, OfferSlot)
    End With
End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(Userindex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(Userindex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(Userindex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(Userindex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), FontTypeNames.FONTTYPE_GUILD))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(Userindex, guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(Userindex, "Propuesta de paz enviada", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim proposal As String
        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(Userindex, guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(Userindex, "Propuesta de alianza enviada", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim details As String
        
        guild = buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(Userindex, guild, RELACIONES_GUILD.ALIADOS, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(Userindex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim details As String
        
        guild = buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(Userindex, guild, RELACIONES_GUILD.PAZ, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(Userindex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim User As String
        Dim details As String
        
        User = buffer.ReadASCIIString()
        
        details = modGuilds.a_DetallesAspirante(Userindex, User)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(Userindex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(Userindex, details)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    Call WriteAlianceProposalsList(Userindex, r_ListaDePropuestas(Userindex, RELACIONES_GUILD.ALIADOS))
End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    Call WritePeaceProposalsList(Userindex, r_ListaDePropuestas(Userindex, RELACIONES_GUILD.PAZ))
End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim errorStr As String
        Dim otherGuildIndex As Integer
        
        guild = buffer.ReadASCIIString()
        
        otherGuildIndex = modGuilds.r_DeclararGuerra(Userindex, guild, errorStr)
        
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarWebSite(Userindex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not modGuilds.a_AceptarAspirante(Userindex, UserName, errorStr) Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
                Call RefreshCharStatus(tUser)
            End If
            
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String
        Dim UserName As String
        Dim Reason As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        
        If Not modGuilds.a_RechazarAspirante(Userindex, UserName, errorStr) Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, errorStr & " : " & Reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, Reason)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        
        GuildIndex = modGuilds.m_EcharMiembroDeClan(Userindex, UserName)
        
        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            Call WriteConsoleMsg(Userindex, "No puedes expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarNoticias(Userindex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendDetallesPersonaje(Userindex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim error As String
        
        If Not modGuilds.v_AbrirElecciones(Userindex, error) Then
            Call WriteConsoleMsg(Userindex, error, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .Name, FontTypeNames.FONTTYPE_GUILD))
        End If
    End With
End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim application As String
        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        application = buffer.ReadASCIIString()
        
        If Not modGuilds.a_NuevoAspirante(Userindex, guild, application, errorStr) Then
           Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
           Call WriteConsoleMsg(Userindex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & guild & ".", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub
''
' Handles the "ShowGuildNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowGuildNews(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMA
'Last Modification: 05/17/06
'
'***************************************************
    
    With UserList(Userindex)
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call modGuilds.SendGuildNews(Userindex)
    End With
End Sub
''
' Handles the "GuildRequestDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendGuildDetails(Userindex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Online" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim Count As Long
    Dim ListaUsers As String
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        For i = 1 To LastUser
            If LenB(UserList(i).Name) <> 0 Then
                If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then _
                    ListaUsers = ListaUsers & UserList(i).Name & ", "
            End If
        Next i
        If ListaUsers <> "" Then
            ListaUsers = left$(ListaUsers, Len(ListaUsers) - 2)
        End If
        Call WriteConsoleMsg(Userindex, "Número de usuarios: " & NumUsers & vbCrLf & ListaUsers, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ)
'If user is invisible, it automatically becomes
'visible before doing the countdown to exit
'04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
'***************************************************
    Dim tUser As Integer
    Dim isNotVisible As Boolean
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        If .flags.Nadando = True Then
            Call WriteConsoleMsg(Userindex, "Debes volver a tu embarcación para poder volver!.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = Userindex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)
                End If
            End If
            
            Call WriteConsoleMsg(Userindex, "Comercio cancelado. ", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(Userindex)
        End If
        
        Call Cerrar_Usuario(Userindex)
    End With
End Sub

''
' Handles the "GuildLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim GuildIndex As Integer
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(Userindex, .Name)
        
        If GuildIndex > 0 Then
            Call WriteConsoleMsg(Userindex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(Userindex, "Tu no puedes salir de ningún clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim earnings As Integer
    Dim percentage As Integer
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenes que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case Npclist(.flags.TargetNPC).NpcType
            Case eNPCType.Banquero
                Call WriteChatOverHead(Userindex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Case eNPCType.Timbero
                If Not .flags.Privilegios And PlayerType.User Then
                    earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
                    If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                        percentage = Int(earnings * 100 / Apuestas.Ganancias)
                    End If
                    
                    If earnings < 0 And Apuestas.Perdidas <> 0 Then
                        percentage = Int(earnings * 100 / Apuestas.Perdidas)
                    End If
                    
                    Call WriteConsoleMsg(Userindex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenás que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's his pet
        If Npclist(.flags.TargetNPC).MaestroUser <> Userindex Then Exit Sub
        
        'Do it!
        Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO
        
        Call Expresar(.flags.TargetNPC, Userindex)
    End With
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenás que seleccionar un personaje, hace click izquierdo sobre ál.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> Userindex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, Userindex)
    End With
End Sub

''
' Handles the "TrainList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's the trainer
        If Npclist(.flags.TargetNPC).NpcType <> eNPCType.Entrenador Then Exit Sub
        
        Call WriteTrainerCreatureList(Userindex, .flags.TargetNPC)
    End With
End Sub

''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!! Solo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(Userindex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(Userindex, "Te acomodás junto a la fogata y comenzás a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.Descansar = Not .flags.Descansar
        Else
            If .flags.Descansar Then
                Call WriteRestOK(Userindex)
                Call WriteConsoleMsg(Userindex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
                .flags.Descansar = False
                Exit Sub
            End If
            
            Call WriteConsoleMsg(Userindex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleCuentaRegresiva(ByVal Userindex As Integer)
 
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
   
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
       
        'Remove packet ID
        Call buffer.ReadByte
       
        Dim Seconds As Byte
       
        Seconds = buffer.ReadByte()
        CuentaRegresivaTimer = Seconds
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Empezando cuenta regresiva desde: " & Seconds & " segundos...!", FontTypeNames.FONTTYPE_GUILD))
       
       
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
   
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
   
    'Destroy auxiliar buffer
    Set buffer = Nothing
   
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/08 (NicoNZ)
'Arreglé un bug que mandaba un index de la meditacion diferente
'al que decia el server.
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡Estás muerto! Solo puedes meditar cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Embarcado Then
            Call WriteConsoleMsg(Userindex, "No se puede meditar dentro del barco.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Can he meditate?
        If .Stats.MaxMAN = 0 Then
             Call WriteConsoleMsg(Userindex, "Sólo las clases mágicas conocen el arte de la meditación", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        'Admins don't have to wait :D
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMAN = .Stats.MaxMAN
            Call WriteConsoleMsg(Userindex, "Mana restaurado", FontTypeNames.FONTTYPE_VENENO)
            Call WriteUpdateMana(Userindex)
            Exit Sub
        End If
        
        Call WriteMeditateToggle(Userindex)
        
        If .flags.Meditando Then _
           Call WriteConsoleMsg(Userindex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        
        .flags.Meditando = Not .flags.Meditando
        
        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            .Counters.tInicioMeditar = (GetTickCount() And &H7FFFFFFF)
            
            Call WriteConsoleMsg(Userindex, "Te estás concentrando. En " & Fix(TIEMPO_INICIOMEDITAR / 1000) & " segundos comenzarás a meditar.", FontTypeNames.FONTTYPE_INFO)
            
            .Char.Loops = INFINITE_LOOPS
            
            'Show proper FX according to level
            If .Stats.ELV < 15 Then
                .Char.FX = FXIDs.FXMEDITARCHICO
            
            ElseIf .Stats.ELV < 30 Then
                .Char.FX = FXIDs.FXMEDITARMEDIANO
            
            ElseIf .Stats.ELV < 40 Then
                .Char.FX = FXIDs.FXMEDITARGRANDE
            
            ElseIf .Stats.ELV < 45 Then
                .Char.FX = FXIDs.FXMEDITARXGRANDE
            
            Else
                .Char.FX = FXIDs.FXMEDITARXXGRANDE
            End If
            
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
        Else
            .Counters.bPuedeMeditar = False
            
            .Char.FX = 0
            .Char.Loops = 0
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        End If
    End With
End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNPC).NpcType <> eNPCType.Revividor _
            And (Npclist(.flags.TargetNPC).NpcType <> eNPCType.ResucitadorNewbie Or Not EsNewbie(Userindex))) _
            Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call RevivirUsuario(Userindex)
        Call WriteConsoleMsg(Userindex, "¡¡Hás sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If (Npclist(.flags.TargetNPC).NpcType <> eNPCType.Revividor _
            And Npclist(.flags.TargetNPC).NpcType <> eNPCType.ResucitadorNewbie) _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Stats.MinHP = .Stats.MaxHP
        
        Call WriteUpdateHP(Userindex)
        
        Call WriteConsoleMsg(Userindex, "¡¡Hás sido curado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    Call SendUserStatsTxt(Userindex, Userindex)
End Sub

''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    Call SendHelp(Userindex)
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
Dim i As Integer
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(Userindex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            'Does the NPC want to trade??
            If Npclist(.flags.TargetNPC).Comercia = 0 Then
                If LenB(Npclist(.flags.TargetNPC).desc) <> 0 Then
                    Call WriteChatOverHead(Userindex, "No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                End If
                
                Exit Sub
            End If
            
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Start commerce....
            Call IniciarComercioNPC(Userindex)
        '[Alejo]
        ElseIf .flags.TargetUser > 0 Then
            'User commerce...
            'Can he commerce??
            If .flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(Userindex, "No puedes vender ítems.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            
            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(Userindex, "¡¡No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is it me??
            If .flags.TargetUser = Userindex Then
                Call WriteConsoleMsg(Userindex, "¡¡No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando = True And _
                UserList(.flags.TargetUser).ComUsu.DestUsu <> Userindex Then
                Call WriteConsoleMsg(Userindex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).Name
            For i = 1 To MAX_OFFER_SLOTS
                .ComUsu.Cant(i) = 0
                .ComUsu.Objeto(i) = 0
            Next i
            .ComUsu.GoldAmount = 0
            
            .ComUsu.Acepto = False
            .ComUsu.Confirmo = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(Userindex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(Userindex, "Primero has click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(Userindex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 4 Then
                Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'If it's the banker....
            If Npclist(.flags.TargetNPC).NpcType = eNPCType.Banquero Then
                Call IniciarDeposito(Userindex)
            End If
        Else
            Call WriteConsoleMsg(Userindex, "Primero has click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComandosVarios(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Tipo As Byte
        Dim Opcion As Byte
        Tipo = .incomingData.ReadByte
        Opcion = .incomingData.ReadByte
        
        If Tipo = 1 Then
            Call HandleEnlistar(Userindex)
        ElseIf Tipo = 2 Then
            Call HandleProteger(Userindex, Opcion)
        End If
    End With
End Sub
Private Sub HandleEnlistar(ByVal Userindex As Integer)
With UserList(Userindex)
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NpcType <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(Userindex, "Debes acercarte más.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(Userindex)
        Else
            Call EnlistarCaos(Userindex)
        End If
End With
End Sub
''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NpcType <> eNPCType.Noble _
                Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
             If .Faccion.ArmadaReal = 0 Then
                 Call WriteChatOverHead(Userindex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call WriteChatOverHead(Userindex, "Tu deber es combatir criminales, cada 100 criminales que derrotes te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
             If .Faccion.FuerzasCaos = 0 Then
                 Call WriteChatOverHead(Userindex, "No perteneces a la legión oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call WriteChatOverHead(Userindex, "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te daré una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NpcType <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
             If .Faccion.ArmadaReal = 0 Then
                 Call WriteChatOverHead(Userindex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call RecompensaArmadaReal(Userindex)
        Else
             If .Faccion.FuerzasCaos = 0 Then
                 Call WriteChatOverHead(Userindex, "No perteneces a la legión oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                 Exit Sub
             End If
             Call RecompensaCaos(Userindex)
        End If
    End With
End Sub

''
' Handles the "RequestMOTD" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMOTD(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    Call SendMOTD(Userindex)
End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/08
'01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    Dim time As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (time Mod 60) & " segundos."
    time = time \ 60
    
    UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
    time = time \ 60
    
    UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
    time = time \ 24
    
    If time = 1 Then
        UpTimeStr = time & " día, " & UpTimeStr
    Else
        UpTimeStr = time & " días, " & UpTimeStr
    End If
    
    Call WriteConsoleMsg(Userindex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub

''
' Handles the "PartyLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyLeave(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    Call mdParty.SalirDeParty(Userindex)
End Sub

''
' Handles the "PartyCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyCreate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    If Not mdParty.PuedeCrearParty(Userindex) Then Exit Sub
    
    Call mdParty.CrearParty(Userindex)
End Sub

''
' Handles the "PartyJoin" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyJoin(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    Call mdParty.SolicitarIngresoAParty(Userindex)
End Sub

''
' Handles the "Inquiry" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    ConsultaPopular.SendInfoEncuesta (Userindex)
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 02/03/09
'02/03/09: ZaMa - Arreglado un indice mal pasado a la funcion de cartel de clanes overhead.
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .GuildIndex > 0 Then
                'Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & Chat))
                Call WriteGuildChat(Userindex, .Name & "> " & Chat)
                'Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageChatOverHead("[" & Chat & "]", .Char.CharIndex, vbYellow))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "PartyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            Call mdParty.BroadCastParty(Userindex, Chat)
'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
            'Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).Pos.map, "||" & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr(UserList(UserIndex).Char.CharIndex))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "CentinelReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCentinelReport(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call CentinelaCheckClave(Userindex, .incomingData.ReadInteger())
    End With
End Sub

''
' Handles the "GuildOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim onlineList As String
        
        onlineList = modGuilds.m_ListaDeMiembrosOnline(Userindex, .GuildIndex)
        
        If .GuildIndex <> 0 Then
            Call WriteConsoleMsg(Userindex, "Compañeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            Call WriteConsoleMsg(Userindex, "No pertences a ningún clan.", FontTypeNames.FONTTYPE_GUILDMSG)
        End If
    End With
End Sub

''
' Handles the "PartyOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyOnline(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadByte
    
    Call mdParty.OnlineParty(Userindex)
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Chat As String
        
        Chat = buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, Userindex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, Userindex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim request As String
        
        request = buffer.ReadASCIIString()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(Userindex, "Su solicitud ha sido enviada", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.Name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not Ayuda.Existe(.Name) Then
            Call WriteConsoleMsg(Userindex, "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM.", FontTypeNames.FONTTYPE_INFO)
            Call Ayuda.Push(.Name)
        Else
            Call Ayuda.Quitar(.Name)
            Call Ayuda.Push(.Name)
            Call WriteConsoleMsg(Userindex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BugReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBugReport(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Dim N As Integer
        
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim bugReport As String
        
        bugReport = buffer.ReadASCIIString()
        
        N = FreeFile
        Open LogPath & "\BUGs.log" For Append Shared As N
        Print #N, "Usuario:" & .Name & "  Fecha:" & Date & "    Hora:" & time
        Print #N, "BUG:"
        Print #N, bugReport
        Print #N, "########################################################################"
        Close #N
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Description As String
        
        Description = buffer.ReadASCIIString()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedés cambiar la descripción estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Else
            If Not AsciiValidos(Description) Then
                Call WriteConsoleMsg(Userindex, "La descripción tiene caractéres inválidos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .desc = Trim$(Description)
                Call WriteConsoleMsg(Userindex, "La descripción ha cambiado.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim vote As String
        Dim errorStr As String
        
        vote = buffer.ReadASCIIString()
        
        If Not modGuilds.v_UsuarioVota(Userindex, vote, errorStr) Then
            Call WriteConsoleMsg(Userindex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub
Private Sub HandleHome(ByVal Userindex As Integer)
'***************************************************
'Author: Budi
'Creation Date: 06/01/2010
'Last Modification: 05/06/10
'Pato - 05/06/10: Add the Ucase$ to prevent problems.
'***************************************************
With UserList(Userindex)
    Call .incomingData.ReadByte
    If .flags.TargetNpcTipo = eNPCType.Gobernador Then
        Call setHome(Userindex, Npclist(.flags.TargetNPC).Ciudad, .flags.TargetNPC)
    Else
        If .flags.Muerto = 1 Then
            'Si es un mapa común y no está en cana
            If Zonas(.zona).Restringir = 0 And (.Counters.Pena = 0) And .SalaIndex = 0 Then
                If .flags.Traveling = 0 Then
                    'If zona(.Hogar).map <> .Pos.map Then
                        Call userGoHome(Userindex)
                    'Else
                    '    Call WriteConsoleMsg(UserIndex, "Ya te encuentras en tu hogar.", FontTypeNames.FONTTYPE_INFO)
                    'End If
                Else
                    Call WriteGotHome(Userindex, False)
                    .flags.Traveling = 0
                    .Counters.goHome = 0
                End If
            Else
                Call WriteConsoleMsg(Userindex, "No puedes usar este comando aquí.", FontTypeNames.FONTTYPE_FIGHT)
            End If
        Else
            Call WriteConsoleMsg(Userindex, "Debes estar muerto para utilizar este comando.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End With
End Sub
''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Name As String
        Dim Count As Integer
        Dim idPj As Long
        Dim Datos As clsMySQLRecordSet
        Dim i As Integer
        
        Name = buffer.ReadASCIIString()
        
        If LenB(Name) <> 0 Then
            If (InStrB(Name, "\") <> 0) Then
                Name = Replace(Name, "\", "")
            End If
            If (InStrB(Name, "/") <> 0) Then
                Name = Replace(Name, "/", "")
            End If
            If (InStrB(Name, ":") <> 0) Then
                Name = Replace(Name, ":", "")
            End If
            If (InStrB(Name, "|") <> 0) Then
                Name = Replace(Name, "|", "")
            End If
            idPj = IdPersonaje(Name)
            If idPj > 0 Then
                Count = mySQL.SQLQuery("SELECT CONCAT(Razon,' : ',Fecha) as Pena FROM penas WHERE IdPj=" & idPj, Datos)

                If Count = 0 Then
                    Call WriteConsoleMsg(Userindex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                Else
                    For i = 1 To Count
                        Call WriteConsoleMsg(Userindex, i & " - " & Datos("Pena"), FontTypeNames.FONTTYPE_INFO)
                    Next i
                    Datos.MoveNext

                End If
            Else
                Call WriteConsoleMsg(Userindex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChangePassword" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangePassword(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Creation Date: 10/10/07
'Last Modified By: Rapsodius
'***************************************************
    If UserList(Userindex).incomingData.Length < 65 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        Dim oldPass As String
        Dim newPass As String
        Dim oldPass2 As String
        
        'Remove packet ID
        Call buffer.ReadByte

        oldPass = buffer.ReadASCIIStringFixed(32)
        newPass = buffer.ReadASCIIStringFixed(32)

        
        If LenB(newPass) = 0 Then
            Call WriteConsoleMsg(Userindex, "Debe especificar una contraseña nueva, inténtelo de nuevo", FontTypeNames.FONTTYPE_INFO)
        Else
            oldPass2 = GetByCampo("SELECT Password FROM cuentas WHERE Id=" & UserList(Userindex).MySQLIdCuenta, "Password")
            
            If oldPass2 <> oldPass Then
                Call WriteConsoleMsg(Userindex, "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call modMySQL.Execute("UPDATE cuentas SET Password=" & Comillas(newPass) & " WHERE Id=" & UserList(Userindex).MySQLIdCuenta)
                Call WriteConsoleMsg(Userindex, "La contraseña fue cambiada con éxito", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Integer
        
        Amount = .incomingData.ReadInteger()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
        ElseIf .flags.TargetNPC = 0 Then
            'Validate target NPC
            Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, has click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Npclist(.flags.TargetNPC).NpcType <> eNPCType.Timbero Then
            Call WriteChatOverHead(Userindex, "No tengo ningún interés en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(Userindex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf Amount > 5000 Then
            Call WriteChatOverHead(Userindex, "El máximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf .Stats.GLD < Amount Then
            Call WriteChatOverHead(Userindex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            If RandomNumber(1, 100) <= 47 Then
                .Stats.GLD = .Stats.GLD + Amount
                Call WriteChatOverHead(Userindex, "Felicidades! Has ganado " & CStr(Amount) & " monedas de oro!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.GLD = .Stats.GLD - Amount
                Call WriteChatOverHead(Userindex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(Userindex)
        End If
    End With
End Sub

''
' Handles the "InquiryVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiryVote(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim opt As Byte
        
        opt = .incomingData.ReadByte()
        
        Call WriteConsoleMsg(Userindex, ConsultaPopular.doVotar(Userindex, opt), FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, has click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NpcType <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Amount > 0 And Amount <= .Stats.Banco Then
             .Stats.Banco = .Stats.Banco - Amount
             .Stats.GLD = .Stats.GLD + Amount
             Call WriteChatOverHead(Userindex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
             Call WriteChatOverHead(Userindex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
        
        Call WriteUpdateGold(Userindex)
        Call WriteUpdateBankGold(Userindex)
    End With
End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, has click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NpcType = eNPCType.Noble Then
           'Quit the Royal Army?
           If .Faccion.ArmadaReal = 1 Then
               If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
                   Call ExpulsarFaccionReal(Userindex)
                   Call WriteChatOverHead(Userindex, "Serás bienvenido a las fuerzas imperiales si deseas regresar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
               Else
                   Call WriteChatOverHead(Userindex, "¡¡¡Sal de aquí bufón!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
               End If
            'Quit the Chaos Legion??
           ElseIf .Faccion.FuerzasCaos = 1 Then
               If Npclist(.flags.TargetNPC).flags.Faccion = 1 Then
                   Call ExpulsarFaccionCaos(Userindex, False)
                   Call WriteChatOverHead(Userindex, "Ya volverás arrastrandote.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
               Else
                   Call WriteChatOverHead(Userindex, "Sal de aquí maldito criminal", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
               End If
           Else
               Call WriteChatOverHead(Userindex, "¡No perteneces a ninguna facción!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
           End If
        End If
    End With
End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, has click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NpcType <> eNPCType.Banquero Then Exit Sub
        
        If Amount > 0 And Amount <= .Stats.GLD Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.GLD = .Stats.GLD - Amount
            Call WriteChatOverHead(Userindex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Call WriteUpdateGold(Userindex)
            Call WriteUpdateBankGold(Userindex)
        Else
            Call WriteChatOverHead(Userindex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Text As String
        
        Text = buffer.ReadASCIIString()
        
        If .flags.Silenciado = 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Text)
            
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(LCase$(.Name) & " DENUNCIA: " & Text, FontTypeNames.FONTTYPE_GUILDMSG))
            Call WriteConsoleMsg(Userindex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildFundate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim clanType As eClanType
        Dim error As String
        
        clanType = .incomingData.ReadByte()
        
        Select Case UCase$(Trim(clanType))
            Case eClanType.ct_RoyalArmy
                .FundandoGuildAlineacion = ALINEACION_ARMADA
            Case eClanType.ct_Evil
                .FundandoGuildAlineacion = ALINEACION_LEGION
            Case eClanType.ct_Neutral
                .FundandoGuildAlineacion = ALINEACION_NEUTRO
            Case eClanType.ct_GM
                .FundandoGuildAlineacion = ALINEACION_MASTER
            Case eClanType.ct_Legal
                .FundandoGuildAlineacion = ALINEACION_CIUDA
            Case eClanType.ct_Criminal
                .FundandoGuildAlineacion = ALINEACION_CRIMINAL
            Case Else
                Call WriteConsoleMsg(Userindex, "Alineación inválida.", FontTypeNames.FONTTYPE_GUILD)
                Exit Sub
        End Select
        
        If modGuilds.PuedeFundarUnClan(Userindex, .FundandoGuildAlineacion, error) Then
            Call WriteShowGuildFundationForm(Userindex)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(Userindex, error, FontTypeNames.FONTTYPE_GUILD)
        End If
    End With
End Sub

''
' Handles the "PartyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyKick(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/05/09
'Last Modification by: Marco Vanotti (Marco)
'- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If UserPuedeEjecutarComandos(Userindex) Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call mdParty.ExpulsarDeParty(Userindex, tUser)
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                
                Call WriteConsoleMsg(Userindex, LCase(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "PartySetLeader" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartySetLeader(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/05/09
'Last Modification by: Marco Vanotti (MarKoxX)
'- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
'on error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = buffer.ReadASCIIString()
        If UserPuedeEjecutarComandos(Userindex) Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(UserName) And rank) <= (.flags.Privilegios And rank) Then
                    Call mdParty.TransformarEnLider(Userindex, tUser)
                Else
                    Call WriteConsoleMsg(Userindex, LCase(UserList(tUser).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
                End If
                
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                Call WriteConsoleMsg(Userindex, LCase(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "PartyAcceptMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyAcceptMember(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/05/09
'Last Modification by: Marco Vanotti (Marco)
'- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        Dim bUserVivo As Boolean
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = buffer.ReadASCIIString()
        If UserList(Userindex).flags.Muerto Then
            Call WriteConsoleMsg(Userindex, "¡Estás muerto!", FontTypeNames.FONTTYPE_PARTY)
        Else
            bUserVivo = True
        End If
        
        If mdParty.UserPuedeEjecutarComandos(Userindex) And bUserVivo Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                'Validate administrative ranks - don't allow users to spoof online GMs
                If (UserList(tUser).flags.Privilegios And rank) <= (.flags.Privilegios And rank) Then
                    Call mdParty.AprobarIngresoAParty(Userindex, tUser)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes incorporar a tu party a personajes de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                End If
                
                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(UserName) And rank) <= (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(Userindex, LCase(UserName) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes incorporar a tu party a personajes de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GuildMemberList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        Dim memberCount As Integer
        Dim i As Long
        Dim UserName As String
        
        guild = buffer.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(guild, "\") <> 0) Then
                guild = Replace(guild, "\", "")
            End If
            If (InStrB(guild, "/") <> 0) Then
                guild = Replace(guild, "/", "")
            End If
            
            If Not FileExist(App.Path & "\guilds\" & guild & "-members.mem") Then
                Call WriteConsoleMsg(Userindex, "No existe el clan: " & guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To memberCount
                    UserName = GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(Userindex, UserName & "<" & guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuild(ByVal Userindex As Integer)
Dim packetID As Byte
'Remove packet ID
UserList(Userindex).incomingData.ReadByte
    
packetID = UserList(Userindex).incomingData.PeekByte()
Select Case packetID
        Case ClientPacketIDGuild.GuildAcceptPeace        'ACEPPEAT
            Call HandleGuildAcceptPeace(Userindex)
        
        Case ClientPacketIDGuild.GuildRejectAlliance     'RECPALIA
            Call HandleGuildRejectAlliance(Userindex)
        
        Case ClientPacketIDGuild.GuildRejectPeace        'RECPPEAT
            Call HandleGuildRejectPeace(Userindex)
        
        Case ClientPacketIDGuild.GuildAcceptAlliance     'ACEPALIA
            Call HandleGuildAcceptAlliance(Userindex)
        
        Case ClientPacketIDGuild.GuildOfferPeace         'PEACEOFF
            Call HandleGuildOfferPeace(Userindex)
        
        Case ClientPacketIDGuild.GuildOfferAlliance      'ALLIEOFF
            Call HandleGuildOfferAlliance(Userindex)
        
        Case ClientPacketIDGuild.GuildAllianceDetails    'ALLIEDET
            Call HandleGuildAllianceDetails(Userindex)
        
        Case ClientPacketIDGuild.GuildPeaceDetails       'PEACEDET
            Call HandleGuildPeaceDetails(Userindex)
        
        Case ClientPacketIDGuild.GuildRequestJoinerInfo  'ENVCOMEN
            Call HandleGuildRequestJoinerInfo(Userindex)
        
        Case ClientPacketIDGuild.GuildAlliancePropList   'ENVALPRO
            Call HandleGuildAlliancePropList(Userindex)
        
        Case ClientPacketIDGuild.GuildPeacePropList      'ENVPROPP
            Call HandleGuildPeacePropList(Userindex)
        
        Case ClientPacketIDGuild.GuildDeclareWar         'DECGUERR
            Call HandleGuildDeclareWar(Userindex)
        
        Case ClientPacketIDGuild.GuildNewWebsite         'NEWWEBSI
            Call HandleGuildNewWebsite(Userindex)
        
        Case ClientPacketIDGuild.GuildAcceptNewMember    'ACEPTARI
            Call HandleGuildAcceptNewMember(Userindex)
        
        Case ClientPacketIDGuild.GuildRejectNewMember    'RECHAZAR
            Call HandleGuildRejectNewMember(Userindex)
        
        Case ClientPacketIDGuild.GuildKickMember         'ECHARCLA
            Call HandleGuildKickMember(Userindex)
        
        Case ClientPacketIDGuild.GuildUpdateNews         'ACTGNEWS
            Call HandleGuildUpdateNews(Userindex)
        
        Case ClientPacketIDGuild.GuildMemberInfo         '1HRINFO<
            Call HandleGuildMemberInfo(Userindex)
        
        Case ClientPacketIDGuild.GuildOpenElections      'ABREELEC
            Call HandleGuildOpenElections(Userindex)
        
        Case ClientPacketIDGuild.GuildRequestMembership  'SOLICITUD
            Call HandleGuildRequestMembership(Userindex)
        
        Case ClientPacketIDGuild.GuildRequestDetails     'CLANDETAILS
            Call HandleGuildRequestDetails(Userindex)
            
        Case ClientPacketIDGuild.ShowGuildNews
            Call HandleShowGuildNews(Userindex)
        Case Else
            'ERROR : Abort!
            Call CloseSocket(Userindex)
End Select
End Sub

Private Sub HandleGM(ByVal Userindex As Integer)
Dim packetID As Byte
'Remove packet ID
packetID = UserList(Userindex).incomingData.ReadByte()
packetID = UserList(Userindex).incomingData.PeekByte()
Select Case packetID

        Case ClientPacketIDGM.GMMessage               '/GMSG
            Call HandleGMMessage(Userindex)
        
        Case ClientPacketIDGM.showName                '/SHOWNAME
            Call HandleShowName(Userindex)
        
        Case ClientPacketIDGM.OnlineRoyalArmy         '/ONLINEREAL
            Call HandleOnlineRoyalArmy(Userindex)
        
        Case ClientPacketIDGM.onlineChaosLegion       '/ONLINECAOS
            Call HandleOnlineChaosLegion(Userindex)
        
        Case ClientPacketIDGM.GoNearby                '/IRCERCA
            Call HandleGoNearby(Userindex)
        
        Case ClientPacketIDGM.comment                 '/REM
            Call HandleComment(Userindex)
        
        Case ClientPacketIDGM.serverTime              '/HORA
            Call HandleServerTime(Userindex)
        
        Case ClientPacketIDGM.Where                   '/DONDE
            Call HandleWhere(Userindex)
        
        Case ClientPacketIDGM.CreaturesInMap          '/NENE
            Call HandleCreaturesInMap(Userindex)
        
        Case ClientPacketIDGM.WarpChar                '/TELEP
            Call HandleWarpChar(Userindex)
        
        Case ClientPacketIDGM.Silence                 '/SILENCIAR
            Call HandleSilence(Userindex)
        
        Case ClientPacketIDGM.SOSShowList             '/SHOW SOS
            Call HandleSOSShowList(Userindex)
        
        Case ClientPacketIDGM.SOSRemove               'SOSDONE
            Call HandleSOSRemove(Userindex)
        
        Case ClientPacketIDGM.GoToChar                '/IRA
            Call HandleGoToChar(Userindex)
        
        Case ClientPacketIDGM.invisible               '/INVISIBLE
            Call HandleInvisible(Userindex)
        
        Case ClientPacketIDGM.GMPanel                 '/PANELGM
            Call HandleGMPanel(Userindex)
        
        Case ClientPacketIDGM.RequestUserList         'LISTUSU
            Call HandleRequestUserList(Userindex)
        
        Case ClientPacketIDGM.Working                 '/TRABAJANDO
            Call HandleWorking(Userindex)
        
        Case ClientPacketIDGM.Hiding                  '/OCULTANDO
            Call HandleHiding(Userindex)
        
        Case ClientPacketIDGM.Jail                    '/CARCEL
            Call HandleJail(Userindex)
        
        Case ClientPacketIDGM.KillNPC                 '/RMATA
            Call HandleKillNPC(Userindex)
        
        Case ClientPacketIDGM.WarnUser                '/ADVERTENCIA
            Call HandleWarnUser(Userindex)
        
        Case ClientPacketIDGM.EditChar                '/MOD
            Call HandleEditChar(Userindex)
            
        Case ClientPacketIDGM.RequestCharInfo         '/INFO
            Call HandleRequestCharInfo(Userindex)
        
        Case ClientPacketIDGM.RequestCharStats        '/STAT
            Call HandleRequestCharStats(Userindex)
            
        Case ClientPacketIDGM.RequestCharGold         '/BAL
            Call HandleRequestCharGold(Userindex)
            
        Case ClientPacketIDGM.RequestCharInventory    '/INV
            Call HandleRequestCharInventory(Userindex)
            
        Case ClientPacketIDGM.RequestCharBank         '/BOV
            Call HandleRequestCharBank(Userindex)
        
        Case ClientPacketIDGM.RequestCharSkills       '/SKILLS
            Call HandleRequestCharSkills(Userindex)
        
        Case ClientPacketIDGM.ReviveChar              '/REVIVIR
            Call HandleReviveChar(Userindex)
        
        Case ClientPacketIDGM.OnlineGM                '/ONLINEGM
            Call HandleOnlineGM(Userindex)
        
        Case ClientPacketIDGM.OnlineMap               '/ONLINEMAP
            Call HandleOnlineMap(Userindex)
        
        Case ClientPacketIDGM.Forgive                 '/PERDON
            Call HandleForgive(Userindex)
            
        Case ClientPacketIDGM.Kick                    '/ECHAR
            Call HandleKick(Userindex)
            
        Case ClientPacketIDGM.GuildMemberList         '/MIEMBROSCLAN
            Call HandleGuildMemberList(Userindex)
            
        Case ClientPacketIDGM.Execute                 '/EJECUTAR
            Call HandleExecute(Userindex)
            
        Case ClientPacketIDGM.BanChar                 '/BAN
            Call HandleBanChar(Userindex)
            
        Case ClientPacketIDGM.UnbanChar               '/UNBAN
            Call HandleUnbanChar(Userindex)
            
        Case ClientPacketIDGM.NPCFollow               '/SEGUIR
            Call HandleNPCFollow(Userindex)
            
        Case ClientPacketIDGM.SummonChar              '/SUM
            Call HandleSummonChar(Userindex)
            
        Case ClientPacketIDGM.SpawnListRequest        '/CC
            Call HandleSpawnListRequest(Userindex)
            
        Case ClientPacketIDGM.SpawnCreature           'SPA
            Call HandleSpawnCreature(Userindex)
            
        Case ClientPacketIDGM.ResetNPCInventory       '/RESETINV
            Call HandleResetNPCInventory(Userindex)
            
        Case ClientPacketIDGM.CleanWorld              '/LIMPIAR
            Call HandleCleanWorld(Userindex)
            
        Case ClientPacketIDGM.ServerMessage           '/RMSG
            Call HandleServerMessage(Userindex)
            
        Case ClientPacketIDGM.NickToIP                '/NICK2IP
            Call HandleNickToIP(Userindex)
        
        Case ClientPacketIDGM.IPToNick                '/IP2NICK
            Call HandleIPToNick(Userindex)
            
        Case ClientPacketIDGM.GuildOnlineMembers      '/ONCLAN
            Call HandleGuildOnlineMembers(Userindex)
        
        Case ClientPacketIDGM.TeleportCreate          '/CT
            Call HandleTeleportCreate(Userindex)
            
        Case ClientPacketIDGM.TeleportDestroy         '/DT
            Call HandleTeleportDestroy(Userindex)
            
        Case ClientPacketIDGM.RainToggle              '/LLUVIA
            Call HandleRainToggle(Userindex)
        
        Case ClientPacketIDGM.SetCharDescription      '/SETDESC
            Call HandleSetCharDescription(Userindex)
        
        Case ClientPacketIDGM.ForceMIDIToMap          '/FORCEMIDIMAP
            Call HanldeForceMIDIToMap(Userindex)
            
        Case ClientPacketIDGM.ForceWAVEToMap          '/FORCEWAVMAP
            Call HandleForceWAVEToMap(Userindex)
            
        Case ClientPacketIDGM.RoyalArmyMessage        '/REALMSG
            Call HandleRoyalArmyMessage(Userindex)
                        
        Case ClientPacketIDGM.ChaosLegionMessage      '/CAOSMSG
            Call HandleChaosLegionMessage(Userindex)
            
        Case ClientPacketIDGM.CitizenMessage          '/CIUMSG
            Call HandleCitizenMessage(Userindex)
            
        Case ClientPacketIDGM.CriminalMessage         '/CRIMSG
            Call HandleCriminalMessage(Userindex)
            
        Case ClientPacketIDGM.TalkAsNPC               '/TALKAS
            Call HandleTalkAsNPC(Userindex)
        
        Case ClientPacketIDGM.DestroyAllItemsInArea   '/MASSDEST
            Call HandleDestroyAllItemsInArea(Userindex)
            
        Case ClientPacketIDGM.AcceptRoyalCouncilMember '/ACEPTCONSE
            Call HandleAcceptRoyalCouncilMember(Userindex)
            
        Case ClientPacketIDGM.AcceptChaosCouncilMember '/ACEPTCONSECAOS
            Call HandleAcceptChaosCouncilMember(Userindex)
            
        Case ClientPacketIDGM.ItemsInTheFloor         '/PISO
            Call HandleItemsInTheFloor(Userindex)
            
        Case ClientPacketIDGM.MakeDumb                '/ESTUPIDO
            Call HandleMakeDumb(Userindex)
            
        Case ClientPacketIDGM.MakeDumbNoMore          '/NOESTUPIDO
            Call HandleMakeDumbNoMore(Userindex)
            
        Case ClientPacketIDGM.DumpIPTables            '/DUMPSECURITY"
            Call HandleDumpIPTables(Userindex)
            
        Case ClientPacketIDGM.CouncilKick             '/KICKCONSE
            Call HandleCouncilKick(Userindex)
        
        Case ClientPacketIDGM.SetTrigger              '/TRIGGER
            Call HandleSetTrigger(Userindex)
        
        Case ClientPacketIDGM.AskTrigger               '/TRIGGER
            Call HandleAskTrigger(Userindex)
            
        Case ClientPacketIDGM.BannedIPList            '/BANIPLIST
            Call HandleBannedIPList(Userindex)
        
        Case ClientPacketIDGM.BannedIPReload          '/BANIPRELOAD
            Call HandleBannedIPReload(Userindex)
        
        Case ClientPacketIDGM.GuildBan                '/BANCLAN
            Call HandleGuildBan(Userindex)
        
        Case ClientPacketIDGM.BanIP                   '/BANIP
            Call HandleBanIP(Userindex)
        
        Case ClientPacketIDGM.UnbanIP                 '/UNBANIP
            Call HandleUnbanIP(Userindex)
        
        Case ClientPacketIDGM.CreateItem              '/CI
            Call HandleCreateItem(Userindex)
        
        Case ClientPacketIDGM.DestroyItems            '/DEST
            Call HandleDestroyItems(Userindex)
        
        Case ClientPacketIDGM.ChaosLegionKick         '/NOCAOS
            Call HandleChaosLegionKick(Userindex)
        
        Case ClientPacketIDGM.RoyalArmyKick           '/NOREAL
            Call HandleRoyalArmyKick(Userindex)
        
        Case ClientPacketIDGM.ForceWAVEAll            '/FORCEWAV
            Call HandleForceWAVEAll(Userindex)
        
        Case ClientPacketIDGM.RemovePunishment        '/BORRARPENA
            Call HandleRemovePunishment(Userindex)
        
        Case ClientPacketIDGM.TileBlockedToggle       '/BLOQ
            Call HandleTileBlockedToggle(Userindex)
        
        Case ClientPacketIDGM.KillNPCNoRespawn        '/MATA
            Call HandleKillNPCNoRespawn(Userindex)
        
        Case ClientPacketIDGM.KillAllNearbyNPCs       '/MASSKILL
            Call HandleKillAllNearbyNPCs(Userindex)
        
        Case ClientPacketIDGM.LastIP                  '/LASTIP
            Call HandleLastIP(Userindex)
        
        Case ClientPacketIDGM.ChangeMOTD              '/MOTDCAMBIA
            Call HandleChangeMOTD(Userindex)
        
        Case ClientPacketIDGM.SetMOTD                 'ZMOTD
            Call HandleSetMOTD(Userindex)
        
        Case ClientPacketIDGM.SystemMessage           '/SMSG
            Call HandleSystemMessage(Userindex)
        
        Case ClientPacketIDGM.createnpc               '/ACC
            Call HandleCreateNPC(Userindex)
        
        Case ClientPacketIDGM.CreateNPCWithRespawn    '/RACC
            Call HandleCreateNPCWithRespawn(Userindex)
        
        Case ClientPacketIDGM.ImperialArmour          '/AI1 - 4
            Call HandleImperialArmour(Userindex)
        
        Case ClientPacketIDGM.ChaosArmour             '/AC1 - 4
            Call HandleChaosArmour(Userindex)
        
        Case ClientPacketIDGM.NavigateToggle          '/NAVE
            Call HandleNavigateToggle(Userindex)
        
        Case ClientPacketIDGM.ServerOpenToUsersToggle '/HABILITAR
            Call HandleServerOpenToUsersToggle(Userindex)
        
        Case ClientPacketIDGM.TurnOffServer           '/APAGAR
            Call HandleTurnOffServer(Userindex)
        
        Case ClientPacketIDGM.TurnCriminal            '/CONDEN
            Call HandleTurnCriminal(Userindex)
        
        Case ClientPacketIDGM.ResetFactions           '/RAJAR
            Call HandleResetFactions(Userindex)
        
        Case ClientPacketIDGM.RemoveCharFromGuild     '/RAJARCLAN
            Call HandleRemoveCharFromGuild(Userindex)
        
        Case ClientPacketIDGM.RequestCharMail         '/LASTEMAIL
            Call HandleRequestCharMail(Userindex)
        
        Case ClientPacketIDGM.AlterPassword           '/APASS
            Call HandleAlterPassword(Userindex)
        
        Case ClientPacketIDGM.AlterMail               '/AEMAIL
            Call HandleAlterMail(Userindex)
        
        Case ClientPacketIDGM.AlterName               '/ANAME
            Call HandleAlterName(Userindex)
        
        Case ClientPacketIDGM.ToggleCentinelActivated '/CENTINELAACTIVADO
            Call HandleToggleCentinelActivated(Userindex)
        
        Case ClientPacketIDGM.DoBackUp                '/DOBACKUP
            Call HandleDoBackUp(Userindex)
        
        Case ClientPacketIDGM.ShowGuildMessages       '/SHOWCMSG
            Call HandleShowGuildMessages(Userindex)
        
        Case ClientPacketIDGM.SaveMap                 '/GUARDAMAPA
            Call HandleSaveMap(Userindex)
                
        Case ClientPacketIDGM.SaveChars               '/GRABAR
            Call HandleSaveChars(Userindex)
        
        Case ClientPacketIDGM.CleanSOS                '/BORRAR SOS
            Call HandleCleanSOS(Userindex)
        
        Case ClientPacketIDGM.ShowServerForm          '/SHOW INT
            Call HandleShowServerForm(Userindex)
        
        Case ClientPacketIDGM.KickAllChars            '/ECHARTODOSPJS
            Call HandleKickAllChars(Userindex)
        
        Case ClientPacketIDGM.ReloadNPCs              '/RELOADNPCS
            Call HandleReloadNPCs(Userindex)
        
        Case ClientPacketIDGM.ReloadServerIni         '/RELOADSINI
            Call HandleReloadServerIni(Userindex)
        
        Case ClientPacketIDGM.ReloadSpells            '/RELOADHECHIZOS
            Call HandleReloadSpells(Userindex)
        
        Case ClientPacketIDGM.ReloadObjects           '/RELOADOBJ
            Call HandleReloadObjects(Userindex)
        
        Case ClientPacketIDGM.Restart                 '/REINICIAR
            Call HandleRestart(Userindex)
        
        Case ClientPacketIDGM.ResetAutoUpdate         '/AUTOUPDATE
            Call HandleResetAutoUpdate(Userindex)
        
        Case ClientPacketIDGM.ChatColor               '/CHATCOLOR
            Call HandleChatColor(Userindex)
        
        Case ClientPacketIDGM.Ignored                 '/IGNORADO
            Call HandleIgnored(Userindex)
        
        Case ClientPacketIDGM.CheckSlot               '/SLOT
            Call HandleCheckSlot(Userindex)
            
        Case ClientPacketIDGM.SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
            Call HandleSetIniVar(Userindex)
            
        Case ClientPacketIDGM.AddGM               '/ADDGM Nombre RANGO
            Call HandleAddGM(Userindex)
            
        Case ClientPacketIDGM.CuentaRegresiva               '/CUENTA TIEMPO
            Call HandleCuentaRegresiva(Userindex)
            
        Case Else
            'ERROR : Abort!
            Call CloseSocket(Userindex)
End Select
End Sub

Private Sub HandleGMMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        
        message = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Mensaje a Gms:" & message)
        
            If LenB(message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(message)
            
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & "> " & message, FontTypeNames.FONTTYPE_GMMSG))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName 'Show / Hide the name
            
            Call RefreshCharStatus(Userindex)
        End If
    End With
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i As Long
        Dim list As String

        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.ArmadaReal = 1 Then
                    If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Or _
                      .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).Name & ", "
                    End If
                End If
            End If
        Next i
    End With
    
    If Len(list) > 0 Then
        Call WriteConsoleMsg(Userindex, "Armadas conectados: " & left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(Userindex, "No hay Armadas conectados", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i As Long
        Dim list As String

        For i = 1 To LastUser
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.FuerzasCaos = 1 Then
                    If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Or _
                      .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).Name & ", "
                    End If
                End If
            End If
        Next i
    End With

    If Len(list) > 0 Then
        Call WriteConsoleMsg(Userindex, "Caos conectados: " & left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(Userindex, "No hay Caos conectados", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/07
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString()
        
        Dim tIndex As Integer
        Dim X As Long
        Dim Y As Long
        Dim i As Long
        Dim found As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i
                                If MapData(UserList(tIndex).Pos.map).Tile(X, Y).Userindex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.map, X, Y, True, True) Then
                                        Call WarpUserChar(Userindex, UserList(tIndex).Pos.map, X, Y, True)
                                        Call LogGM(.Name, "/IRCERCA " & UserName & " Mapa:" & UserList(tIndex).Pos.map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
                                        found = True
                                        Exit For
                                    End If
                                End If
                            Next Y
                            
                            If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next X
                        
                        If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i
                    
                    'No space found??
                    If Not found Then
                        Call WriteConsoleMsg(Userindex, "Todos los lugares están ocupados.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim comment As String
        comment = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Comentario: " & comment)
            Call WriteConsoleMsg(Userindex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Call LogGM(.Name, "Hora.")
    End With
    
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & time & " " & Date, FontTypeNames.FONTTYPE_INFO))
End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(Userindex, "Ubicación  " & UserName & ": " & UserList(tUser).Pos.map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/Donde " & UserName)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 30/07/06
'Pablo (ToxicWaste): modificaciones generales para simplificar la visualización.
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim map As Integer
        Dim i, j As Long
        Dim NPCcount1, NPCcount2 As Integer
        Dim NPCcant1() As Integer
        Dim NPCcant2() As Integer
        Dim List1() As String
        Dim List2() As String
        Dim AreaIndex As Integer
        
        map = .incomingData.ReadInteger()
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        AreaIndex = AreaUser(Userindex)
        'If MapaValido(map) Then
            For i = 1 To LastNPC
                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If Npclist(i).Area = AreaIndex Then
                    '¿esta vivo?
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else
                            For j = 0 To NPCcount1 - 1
                                If left$(List1(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant1(j) = 1
                            End If
                        End If
                    End If
                ElseIf Npclist(i).zona = .zona Then
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile <> 1 And Npclist(i).Stats.Alineacion <> 2 Then
                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else
                            For j = 0 To NPCcount2 - 1
                                If left$(List2(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant2(j) = 1
                            End If
                        End If
                    End If
                End If
            Next i
            
            Call WriteConsoleMsg(Userindex, "Npcs Hostiles en el area " & AreaIndex & " (" & Areas(AreaIndex).X1 & ", " & Areas(AreaIndex).Y1 & ", " & Areas(AreaIndex).X2 & ", " & Areas(AreaIndex).Y2 & "): ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(Userindex, "No hay NPCS Hostiles", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(Userindex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call WriteConsoleMsg(Userindex, "Otros Npcs en la zona: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(Userindex, "No hay más NPCS", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(Userindex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call LogGM(.Name, "Numero enemigos en mapa " & map)
        'End If
    End With
End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/09
'26/03/06: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Integer
        Dim Y As Integer
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        X = .incomingData.ReadInteger
        Y = .incomingData.ReadInteger
        
        Call FindLegalPos(Userindex, .Pos.map, X, Y)
        Call WarpUserChar(Userindex, .Pos.map, X, Y, False, True)
        Call LogGM(.Name, "/TELEPLOC a x:" & .flags.targetX & " Y:" & .flags.targetY & " Map:" & .Pos.map)
    End With
End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(Userindex).incomingData.Length < 7 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim map As Integer
        Dim X As Integer
        Dim Y As Integer
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        map = buffer.ReadInteger()
        X = buffer.ReadInteger()
        Y = buffer.ReadInteger()
        
        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)
                    End If
                Else
                    tUser = Userindex
                End If
            
                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                ElseIf InMapBounds(map, X, Y) Then
                    Call FindLegalPos(tUser, map, X, Y)
                    Call WarpUserChar(tUser, map, X, Y, True, True)
                    Call WriteConsoleMsg(Userindex, UserList(tUser).Name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "Transportó a " & UserList(tUser).Name & " hacia " & "Mapa" & map & " X:" & X & " Y:" & Y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
        
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(Userindex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "ESTIMADO USUARIO, ud ha sido silenciado por los administradores. Sus denuncias serán ignoradas por el servidor de aquí en más. Utilice /GM para contactar un administrador.")
                    Call LogGM(.Name, "/silenciar " & UserList(tUser).Name)
                
                    'Flush the other user's buffer
                    Call FlushBuffer(tUser)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(Userindex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/DESsilenciar " & UserList(tUser).Name)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(Userindex)
    End With
End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then _
            Call Ayuda.Quitar(UserName)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim X As Integer
        Dim Y As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros también lo seamos
            If UserName = "merc1" Then
                X = Npclist(MercaderReal.NpcIndex).Pos.X
                Y = Npclist(MercaderReal.NpcIndex).Pos.Y + 1
                Call FindLegalPos(Userindex, Npclist(MercaderReal.NpcIndex).Pos.map, X, Y)
                Call WarpUserChar(Userindex, Npclist(MercaderReal.NpcIndex).Pos.map, X, Y, True, True)
            ElseIf UserName = "merc2" Then
                X = Npclist(MercaderCaos.NpcIndex).Pos.X
                Y = Npclist(MercaderCaos.NpcIndex).Pos.Y + 1
                Call FindLegalPos(Userindex, Npclist(MercaderReal.NpcIndex).Pos.map, X, Y)
                Call WarpUserChar(Userindex, Npclist(MercaderReal.NpcIndex).Pos.map, X, Y, True, True)
            ElseIf Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    X = UserList(tUser).Pos.X
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(Userindex, UserList(tUser).Pos.map, X, Y)
                    
                    Call WarpUserChar(Userindex, UserList(tUser).Pos.map, X, Y, True, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                        Call FlushBuffer(tUser)
                    End If
                    
                    Call LogGM(.Name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call DoAdminInvisible(Userindex)
        Call LogGM(.Name, "/INVISIBLE")
    End With
End Sub
''
' Handles the "InitCrafting" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInitCrafting(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'
'***************************************************
    Dim TotalItems As Long
    Dim ItemsPorCiclo As Integer
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        TotalItems = .incomingData.ReadLong
        ItemsPorCiclo = .incomingData.ReadInteger
        
        If TotalItems > 0 Then
            
            .Construir.Cantidad = TotalItems
            .Construir.PorCiclo = MinimoInt(MaxItemsConstruibles(Userindex), ItemsPorCiclo)
            
        End If
    End With
End Sub
''
' Handles the "ItemUpgrade" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleItemUpgrade(ByVal Userindex As Integer)
'***************************************************
'Author: Torres Patricio
'Last Modification: 12/09/09
'
'***************************************************
    With UserList(Userindex)
        Dim ItemIndex As Integer
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        ItemIndex = .incomingData.ReadInteger()
        
        If ItemIndex <= 0 Then Exit Sub
        If Not TieneObjetos(ItemIndex, 1, Userindex) Then Exit Sub
        
        Call DoUpgrade(Userindex, ItemIndex)
    End With
End Sub


''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WriteShowGMPanelForm(Userindex)
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'Last modified by: Lucas Tavolaro Ortiz (Tavo)
'I haven`t found a solution to split, so i make an array of names
'***************************************************
    Dim i As Long
    Dim names() As String
    Dim Count As Long
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        ReDim names(1 To LastUser) As String
        Count = 1
        
        For i = 1 To LastUser
            If (LenB(UserList(i).Name) <> 0) Then
                If UserList(i).flags.Privilegios And PlayerType.User Then
                    names(Count) = UserList(i).Name
                    Count = Count + 1
                End If
            End If
        Next i
        
        If Count > 1 Then Call WriteUserNameList(Userindex, names(), Count - 1)
    End With
End Sub

''
' Handles the "Working" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim Users As String
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                Users = Users & ", " & UserList(i).Name
                
                ' Display the user being checked by the centinel
                If modCentinela.Centinela.RevisandoUserIndex = i Then _
                    Users = Users & " (*)"
            End If
        Next i
        
        If LenB(Users) <> 0 Then
            Users = Right$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(Userindex, "Usuarios trabajando: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "No hay usuarios trabajando", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Hiding" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim Users As String
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            If (LenB(UserList(i).Name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                Users = Users & UserList(i).Name & ", "
            End If
        Next i
        
        If LenB(Users) <> 0 Then
            Users = left$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(Userindex, "Usuarios ocultandose: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "No hay usuarios ocultandose", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Reason As String
        Dim jailTime As Byte
        Dim Count As Byte
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        jailTime = buffer.ReadByte()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
        End If
        
        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(Userindex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteConsoleMsg(Userindex, "No podés encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > 600 Then
                        Call WriteConsoleMsg(Userindex, "No podés encarcelar por más de 600 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                        End If
                        If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                        End If
                        
                        If PersonajeExiste(UserName) Then
                            'ponemos el flag de ban a 1
                            modMySQL.Execute ("UPDATE pjs SET Penas=Penas+1 WHERE Nombre=" & Comillas(UserName))
                            'ponemos la pena
                            modMySQL.Execute ("INSERT INTO penas (IdPj, Razon, Fecha, IdGM,Tiempo) VALUES (" & UserList(tUser).MySQLId & "," & Comillas(Reason) & ",NOW()," & .MySQLId & "," & jailTime & ")")
                    
                        End If
                        
                        Call Encarcelar(tUser, jailTime, .Name)
                        Call LogGM(.Name, " encarcelo a " & UserName)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/22/08 (NicoNZ)
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim tNPC As Integer
        Dim auxNPC As NPC
        
        'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
        If .flags.Privilegios And PlayerType.Consejero Then
            If .zona = ZONA_PRETORIANO Then
                Call WriteConsoleMsg(Userindex, "Los consejeros no pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        tNPC = .flags.TargetNPC
        
        If tNPC > 0 Then
            Call WriteConsoleMsg(Userindex, "RMatas (con posible respawn) a: " & Npclist(tNPC).Name, FontTypeNames.FONTTYPE_INFO)
            
            auxNPC = Npclist(tNPC)
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
            
            .flags.TargetNPC = 0
        Else
            Call WriteConsoleMsg(Userindex, "Debes hacer click sobre el NPC antes", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Reason As String
        Dim privs As PlayerType
        Dim Count As Byte
        Dim idPj As Long
        
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(Userindex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                privs = UserDarPrivilegioLevel(UserName)
                
                If Not privs And PlayerType.User Then
                    Call WriteConsoleMsg(Userindex, "No podés advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                    End If
                    If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                    End If
                    
                    If PersonajeExiste(UserName) Then
                        idPj = val(GetByCampo("SELECT Id FROM pjs WHERE Nombre=" & Comillas(UserName), "Id"))
                        'ponemos el flag de ban a 1
                        modMySQL.Execute ("UPDATE pjs SET Penas=Penas+1 WHERE Id=" & idPj)
                        'ponemos la pena
                        modMySQL.Execute ("INSERT INTO penas (IdPj, Razon, Fecha, IdGM, Tiempo) VALUES (" & idPj & "," & Comillas(.Name & ": ADVERTENCIA por: " & LCase$(Reason)) & ",NOW()," & .MySQLId & ",0)")

                        Call WriteConsoleMsg(Userindex, "Has advertido a " & UCase$(UserName), FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, " advirtio a " & UserName)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "EditChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 11/06/2009
'02/03/2009: ZaMa - Cuando editas nivel, chequea si el pj puede permanecer en clan faccionario
'11/06/2009: ZaMa - Todos los comandos se pueden usar aunque el pj este offline
'***************************************************
    If UserList(Userindex).incomingData.Length < 8 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim Opcion As Byte
        Dim Arg1 As String
        Dim Arg2 As String
        Dim valido As Boolean
        Dim LoopC As Byte
        Dim CommandString As String
        Dim N As Byte
        Dim UserCharPath As String
        Dim Var As Long
        
        
        UserName = Replace(buffer.ReadASCIIString(), "+", " ")
        
        If UCase$(UserName) = "YO" Then
            tUser = Userindex
        Else
            tUser = NameIndex(UserName)
        End If
        
        Opcion = buffer.ReadByte()
        Arg1 = buffer.ReadASCIIString()
        Arg2 = buffer.ReadASCIIString()
        
        If .flags.Privilegios And PlayerType.RoleMaster Then
            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)
                Case PlayerType.Consejero
                    ' Los RMs consejeros sólo se pueden editar su head, body y level
                    valido = tUser = Userindex And _
                            (Opcion = eEditOptions.eo_Body Or Opcion = eEditOptions.eo_Head Or Opcion = eEditOptions.eo_Level)
                
                Case PlayerType.SemiDios
                    ' Los RMs sólo se pueden editar su level y el head y body de cualquiera
                    valido = (Opcion = eEditOptions.eo_Level And tUser = Userindex) _
                            Or Opcion = eEditOptions.eo_Body Or Opcion = eEditOptions.eo_Head
                
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level sólo lo puede hacer sobre sí mismo
                    valido = (Opcion = eEditOptions.eo_Level And tUser = Userindex) Or _
                            Opcion = eEditOptions.eo_Body Or _
                            Opcion = eEditOptions.eo_Head Or _
                            Opcion = eEditOptions.eo_CiticensKilled Or _
                            Opcion = eEditOptions.eo_CriminalsKilled Or _
                            Opcion = eEditOptions.eo_Class Or _
                            Opcion = eEditOptions.eo_Skills Or _
                            Opcion = eEditOptions.eo_addGold
            End Select
            
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then   'Si no es RM debe ser dios para poder usar este comando
            valido = True
        End If

        If valido Then
            If tUser <= 0 And Not PersonajeExiste(UserName) Then
                Call WriteConsoleMsg(Userindex, "Esta intentando editar un usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                Call LogGM(.Name, "Intento editar un usuario inexistente.")
            Else
                'For making the Log
                CommandString = "/MOD "
                
                Select Case Opcion
                    Case eEditOptions.eo_Gold
                        If val(Arg1) < MAX_ORO_EDIT Then
                            If tUser <= 0 Then ' Esta offline?
                                modMySQL.Execute ("UPDATE pjs SET GLD=" & val(Arg1) & " WHERE Nombre=" & Comillas(UserName))
                                Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Stats.GLD = val(Arg1)
                                Call WriteUpdateGold(tUser)
                            End If
                        Else
                            Call WriteConsoleMsg(Userindex, "No esta permitido utilizar valores mayores. Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "ORO "
                
                    Case eEditOptions.eo_Experience
                        If val(Arg1) > 20000000 Then
                                Arg1 = 20000000
                        End If
                        
                        If tUser <= 0 Then ' Offline
                            modMySQL.Execute ("UPDATE pjs SET EXP=EXP+" & val(Arg1) & " WHERE Nombre=" & Comillas(UserName))
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                            Call CheckUserLevel(tUser)
                            Call WriteUpdateExp(tUser)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "EXP "
                    
                    Case eEditOptions.eo_Body
                        If tUser <= 0 Then
                            modMySQL.Execute ("UPDATE pjs SET Body=" & val(Arg1) & " WHERE Nombre=" & Comillas(UserName))
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "BODY "
                    
                    Case eEditOptions.eo_Head
                        If tUser <= 0 Then
                            modMySQL.Execute ("UPDATE pjs SET Head=" & val(Arg1) & " WHERE Nombre=" & Comillas(UserName))
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, UserList(tUser).Char.Body, val(Arg1), UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "HEAD "
                    
                    Case eEditOptions.eo_CriminalsKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            modMySQL.Execute ("UPDATE pjs SET CrimMatados=" & Var & " WHERE Nombre=" & Comillas(UserName))
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Faccion.CriminalesMatados = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CRI "
                    
                    Case eEditOptions.eo_CiticensKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            modMySQL.Execute ("UPDATE pjs SET CiudMatados=" & Var & " WHERE Nombre=" & Comillas(UserName))
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Faccion.CiudadanosMatados = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CIU "
                    
                    Case eEditOptions.eo_Level
                        If val(Arg1) > STAT_MAXELV Then
                            Arg1 = CStr(STAT_MAXELV)
                            Call WriteConsoleMsg(Userindex, "No puedes tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)
                        End If
                        
                        ' Chequeamos si puede permanecer en el clan
                        If val(Arg1) >= 25 Then
                            
                            Dim GI As Integer
                            If tUser <= 0 Then
                                GI = GetByCampo("SELECT GuildIndex FROM pjs WHERE Nombre=" & Comillas(UserName), "GuildIndex")
                            Else
                                GI = UserList(tUser).GuildIndex
                            End If
                            
                            If GI > 0 Then
                                If modGuilds.GuildAlignment(GI) = "Legión oscura" Or modGuilds.GuildAlignment(GI) = "Armada Real" Then
                                    'We get here, so guild has factionary alignment, we have to expulse the user
                                    Call modGuilds.m_EcharMiembroDeClan(-1, UserName)
                                    
                                    Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(UserName & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
                                    ' Si esta online le avisamos
                                    If tUser > 0 Then _
                                        Call WriteConsoleMsg(tUser, "¡Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearás! Por esta razón, hasta tanto no te enlistes en la Facción bajo la cual tu clan está alineado, estarás excluído del mismo.", FontTypeNames.FONTTYPE_GUILD)
                                End If
                            End If
                        End If
                        
                        If tUser <= 0 Then ' Offline
                            modMySQL.Execute ("UPDATE pjs SET ELV=" & val(Arg1) & " WHERE Nombre=" & Comillas(UserName))
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.ELV = val(Arg1)
                            Call WriteUpdateUserStats(tUser)
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "LEVEL "
                    
                    Case eEditOptions.eo_Class
                        For LoopC = 1 To NUMCLASES
                            If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
                        Next LoopC
                            
                        If LoopC > NUMCLASES Then
                            Call WriteConsoleMsg(Userindex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then ' Offline
                                modMySQL.Execute ("UPDATE pjs SET Clase=" & LoopC & " WHERE Nombre=" & Comillas(UserName))
                                Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).clase = LoopC
                            End If
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "CLASE "
                        
                    Case eEditOptions.eo_Skills
                        For LoopC = 1 To NUMSKILLS
                            If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
                        Next LoopC
                        
                        If LoopC > NUMSKILLS Then
                            Call WriteConsoleMsg(Userindex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then ' Offline
                                modMySQL.Execute ("UPDATE pjs SET Sk" & LoopC & "=" & Arg2 & " WHERE Nombre=" & Comillas(UserName))
                                Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)
                            End If
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLS "
                    
                    Case eEditOptions.eo_SkillPointsLeft
                        If tUser <= 0 Then ' Offline
                            modMySQL.Execute ("UPDATE pjs SET SkillPtsLibres=" & val(Arg1) & " WHERE Nombre=" & Comillas(UserName))
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.SkillPts = val(Arg1)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLSLIBRES "
                    
                    Case eEditOptions.eo_Nobleza
                        Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            modMySQL.Execute ("UPDATE pjs SET Rep_Nobles=" & Var & " WHERE Nombre=" & Comillas(UserName))
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Reputacion.NobleRep = Var
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "NOB "
                        
                    Case eEditOptions.eo_Asesino
                        Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            modMySQL.Execute ("UPDATE pjs SET Rep_Asesino=" & Var & " WHERE Nombre=" & Comillas(UserName))
                            Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Reputacion.AsesinoRep = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "ASE "
                    
                    Case eEditOptions.eo_Sex
                        Dim Sex As Byte
                        Sex = IIf(UCase(Arg1) = "MUJER", eGenero.Mujer, 0) ' Mujer?
                        Sex = IIf(UCase(Arg1) = "HOMBRE", eGenero.Hombre, Sex) ' Hombre?
                        
                        If Sex <> 0 Then ' Es Hombre o mujer?
                            If tUser <= 0 Then ' OffLine
                                modMySQL.Execute ("UPDATE pjs SET Genero=" & Sex & " WHERE Nombre=" & Comillas(UserName))
                                Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).genero = Sex
                            End If
                        Else
                            Call WriteConsoleMsg(Userindex, "Genero desconocido. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SEX "
                    
                    Case eEditOptions.eo_Raza
                        Dim raza As Byte
                        
                        Arg1 = UCase$(Arg1)
                        Select Case Arg1
                            Case "HUMANO"
                                raza = eRaza.Humano
                            Case "ELFO"
                                raza = eRaza.Elfo
                            Case "DROW"
                                raza = eRaza.Drow
                            Case "ENANO"
                                raza = eRaza.Enano
                            Case "GNOMO"
                                raza = eRaza.Gnomo
                            Case Else
                                raza = 0
                        End Select
                        
                            
                        If raza = 0 Then
                            Call WriteConsoleMsg(Userindex, "Raza desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                modMySQL.Execute ("UPDATE pjs SET Raza=" & raza & " WHERE Nombre=" & Comillas(UserName))
                                Call WriteConsoleMsg(Userindex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                UserList(tUser).raza = raza
                            End If
                        End If
                            
                        ' Log it
                        CommandString = CommandString & "RAZA "
                        
                    Case eEditOptions.eo_addGold
                    
                        Dim bankGold As Long
                        
                        If Abs(Arg1) > MAX_ORO_EDIT Then
                            Call WriteConsoleMsg(Userindex, "No está permitido utilizar valores mayores a " & MAX_ORO_EDIT & ".", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                modMySQL.Execute ("UPDATE pjs SET Banco=Banco+" & val(Arg1) & " WHERE Nombre=" & Comillas(UserName))
                                Call WriteConsoleMsg(Userindex, "Se le ha agregado " & Arg1 & " monedas de oro a " & UserName & ".", FONTTYPE_TALK)
                            Else
                                UserList(tUser).Stats.Banco = IIf(UserList(tUser).Stats.Banco + val(Arg1) <= 0, 0, UserList(tUser).Stats.Banco + val(Arg1))
                                Call WriteConsoleMsg(tUser, STANDARD_BOUNTY_HUNTER_MESSAGE, FONTTYPE_TALK)
                            End If
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "AGREGAR "
                        
                    Case Else
                        Call WriteConsoleMsg(Userindex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
                        CommandString = CommandString & "UNKOWN "
                        
                End Select
                
                CommandString = CommandString & Arg1 & " " & Arg2
                Call LogGM(.Name, CommandString & " " & UserName)
                
            End If
        End If
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal Userindex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Last Modification by: (liquid).. alto bug zapallo..
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
                
        Dim targetName As String
        Dim targetIndex As Integer
        
        targetName = Replace$(buffer.ReadASCIIString(), "+", " ")
        targetIndex = NameIndex(targetName)
        
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            'is the player offline?
            If targetIndex <= 0 Then
                'don't allow to retrieve administrator's info
                If Not (EsDios(targetName) Or EsAdmin(targetName)) Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline, Buscando en Charfile.", FontTypeNames.FONTTYPE_INFO)
                    Call SendUserStatsTxtOFF(Userindex, targetName)
                End If
            Else
                'don't allow to retrieve administrator's info
                If UserList(targetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(Userindex, targetIndex)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call LogGM(.Name, "/STAT " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline. Leyendo Charfile... ", FontTypeNames.FONTTYPE_INFO)
                
                Call SendUserMiniStatsTxtFromChar(Userindex, UserName)
            Else
                Call SendUserMiniStatsTxt(Userindex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/BAL " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserOROTxtFromChar(Userindex, UserName)
            Else
                Call WriteConsoleMsg(Userindex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco", FontTypeNames.FONTTYPE_TALK)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/INV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserInvTxtFromChar(Userindex, UserName)
            Else
                Call SendUserInvTxt(Userindex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/BOV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserBovedaTxtFromChar(Userindex)
            Else
                Call SendUserBovedaTxt(Userindex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Long
        Dim message As String
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        Dim Datos As clsMySQLRecordSet
        Dim Cant As Long
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/STATS " & UserName)
            
            If tUser <= 0 Then
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                Cant = mySQL.SQLQuery("SELECT SkillPtsLibres,Sk1,Sk2,Sk3,Sk4,Sk5,Sk6,Sk7,Sk8,Sk9,Sk10,Sk11,Sk12,Sk13,Sk14,Sk15,Sk16,Sk17,Sk18,Sk19,Sk20 FROM pjs WHERE Name=" & Comillas(UserName), Datos)
                
                For LoopC = 1 To NUMSKILLS
                    message = message & "CHAR>" & SkillsNames(LoopC) & " = " & Datos("Sk" & LoopC) & vbCrLf
                Next LoopC
                
                Call WriteConsoleMsg(Userindex, message & "CHAR> Libres:" & Datos!SkillPtsLibres, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendUserSkillsTxt(Userindex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = buffer.ReadASCIIString()
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = Userindex
            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                With UserList(tUser)
                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        
                        If .flags.Navegando = 1 Then
                            Call ToggleBoatBody(tUser)
                        Else
                            Call DarCuerpoDesnudo(tUser)
                        End If
                        
                        If .flags.Traveling = 1 Then
                            .flags.Traveling = 0
                            .Counters.goHome = 0
                            Call WriteGotHome(tUser, False)
                        End If
                        
                        
                        Call ChangeUserChar(tUser, .Char.Body, .OrigChar.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        
                        Call WriteConsoleMsg(tUser, UserList(Userindex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(Userindex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                    .Stats.MinHP = .Stats.MaxHP
                End With
                
                Call WriteUpdateHP(tUser)
                
                Call FlushBuffer(tUser)
                
                Call LogGM(.Name, "Resucito a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal Userindex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 12/28/06
'
'***************************************************
    Dim i As Long
    Dim list As String
    Dim priv As PlayerType
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        priv = PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then _
                    list = list & UserList(i).Name & ", "
            End If
        Next i
        
        If LenB(list) <> 0 Then
            list = left$(list, Len(list) - 2)
            Call WriteConsoleMsg(Userindex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 23/03/2009
'23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim map As Integer
        map = .incomingData.ReadInteger
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Dim LoopC As Long
        Dim list As String
        Dim priv As PlayerType
        
        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)
        
        For LoopC = 1 To LastUser
            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.map = map Then
                If UserList(LoopC).flags.Privilegios And priv Then _
                    list = list & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        
        If Len(list) > 2 Then list = left$(list, Len(list) - 2)
        
        Call WriteConsoleMsg(Userindex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Forgive" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForgive(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If EsNewbie(tUser) Then
                    Call VolverCiudadano(tUser)
                Else
                    Call LogGM(.Name, "Intento perdonar un personaje de nivel avanzado.")
                    Call WriteConsoleMsg(Userindex, "Solo se permite perdonar newbies.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(Userindex, "No podes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " echo a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseSocket(tUser)
                    Call LogGM(.Name, "Echo a " & UserName)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(Userindex, "Estás loco?? como vas a piñatear un gm!!!! :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserName, FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.Name, " ejecuto a " & UserName)
                End If
            Else
                Call WriteConsoleMsg(Userindex, "No está online", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim Reason As String
        
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(Userindex, UserName, Reason)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim cantPenas As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            
            If Not PersonajeExiste(UserName) Then
                Call WriteConsoleMsg(Userindex, "Charfile inexistente (no use +)", FontTypeNames.FONTTYPE_INFO)
            Else
                If (val(GetByCampo("SELECT Ban FROM pjs WHERE Nombre='" & UserName & "'", "Ban")) = 1) Then
                    Call UnBan(UserName)
                
                    'penas
                    modMySQL.Execute ("UPDATE pjs SET Penas=Penas+1")
                    modMySQL.Execute ("INSERT INTO penas (IdPj,Razon,Fecha, IdGM, Tiempo) VALUES (" & .MySQLId & ",'UNBAN.',NOW()," & .MySQLId & ",0)")
                
                    Call LogGM(.Name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(Userindex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, UserName & " no esta baneado. Imposible unbanear", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .Name)
            Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
            Npclist(.flags.TargetNPC).flags.Paralizado = 0
            Npclist(.flags.TargetNPC).Contadores.Paralisis = 0
        End If
    End With
End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim X As Integer
        Dim Y As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "El jugador no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or _
                  (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                    X = .Pos.X
                    Y = .Pos.Y + 1
                    Call FindLegalPos(tUser, .Pos.map, X, Y)
                    Call WarpUserChar(tUser, .Pos.map, X, Y, True, True)
                    Call LogGM(.Name, "/SUM " & UserName & " Map:" & .Pos.map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                Else
                    Call WriteConsoleMsg(Userindex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call EnviarSpawnList(Userindex)
    End With
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim NPC As Integer
        Dim SpawnedNpc As Integer
        NPC = .incomingData.ReadInteger()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If NPC > 0 And NPC <= UBound(Declaraciones.SpawnList()) Then
                SpawnedNpc = SpawnNpc(Declaraciones.SpawnList(NPC).NpcIndex, .Pos, True, False, UserList(Userindex).zona)
            End If
            
            Call LogGM(.Name, "Sumoneo " & Declaraciones.SpawnList(NPC).NpcName)
        End If
    End With
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.Name, "/RESETINV " & Npclist(.flags.TargetNPC).Name)
    End With
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LimpiarMundo
    End With
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(message, FontTypeNames.FONTTYPE_TALK))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 24/07/07
'Pablo (ToxicWaste): Agrego para uqe el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim priv As PlayerType
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.Name, "NICK2IP Solicito la IP de " & UserName)

            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                priv = PlayerType.User
            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(Userindex, "El ip de " & UserName & " es " & UserList(tUser).ip, FontTypeNames.FONTTYPE_INFO)
                    Dim ip As String
                    Dim lista As String
                    Dim LoopC As Long
                    ip = UserList(tUser).ip
                    For LoopC = 1 To LastUser
                        If UserList(LoopC).ip = ip Then
                            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).Name & ", "
                                End If
                            End If
                        End If
                    Next LoopC
                    If LenB(lista) <> 0 Then lista = left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(Userindex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(Userindex, "No hay ningun personaje con ese nick", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim ip As String
        Dim LoopC As Long
        Dim lista As String
        Dim priv As PlayerType
        
        ip = .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & ip)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User
        End If

        For LoopC = 1 To LastUser
            If UserList(LoopC).ip = ip Then
                If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).Name & ", "
                    End If
                End If
            End If
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(Userindex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName As String
        Dim tGuild As Integer
        
        GuildName = buffer.ReadASCIIString()
        
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")
        End If
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = GuildIndex(GuildName)
            
            If tGuild > 0 Then
                Call WriteConsoleMsg(Userindex, "Clan " & UCase(GuildName) & ": " & _
                  modGuilds.m_ListaDeMiembrosOnline(Userindex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim mapa As Integer
        Dim X As Integer
        Dim Y As Integer
        
        mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadInteger()
        Y = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.Name, "/CT " & mapa & "," & X & "," & Y)
        
        If Not MapaValido(mapa) Or Not InMapBounds(mapa, X, Y) Then _
            Exit Sub
        
        If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y - 1).TileExit.map > 0 Then _
            Exit Sub
        
        If MapData(mapa).Tile(X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(Userindex, "Hay un objeto en el piso en ese lugar", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapData(mapa).Tile(X, Y).TileExit.map > 0 Then
            Call WriteConsoleMsg(Userindex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim ET As Obj
        ET.Amount = 1
        ET.ObjIndex = 378
        
        Call MakeObj(ET, .Pos.map, .Pos.X, .Pos.Y - 1)
        
        With MapData(.Pos.map).Tile(.Pos.X, .Pos.Y - 1)
            .TileExit.map = mapa
            .TileExit.X = X
            .TileExit.Y = Y
        End With
    End With
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(Userindex)
        Dim mapa As Integer
        Dim X As Integer
        Dim Y As Integer
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        mapa = .flags.TargetMap
        X = .flags.targetX
        Y = .flags.targetY
        
        If Not InMapBounds(mapa, X, Y) Then Exit Sub
        
        With MapData(mapa).Tile(X, Y)
            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.map > 0 Then
                Call LogGM(UserList(Userindex).Name, "/DT: " & mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.Amount, mapa, X, Y)
                
                If MapData(.TileExit.map).Tile(.TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(1, .TileExit.map, .TileExit.X, .TileExit.Y)
                End If
                
                .TileExit.map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
            End If
        End With
    End With
End Sub

''
' Handles the "RainToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRainToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call LogGM(.Name, "/LLUVIA")
        Lloviendo = Not Lloviendo
        CheckFogatasLluvia
        
        If Lloviendo = True And RandomNumber(1, 3) = 1 Then
            LloviendoConTormenta = 1
        Else
            LloviendoConTormenta = 0
        End If
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End With
End Sub

''
' Handles the "SetCharDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim tUser As Integer
        Dim desc As String
        
        desc = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser
            If tUser > 0 Then
                UserList(tUser).DescRM = desc
            Else
                Call WriteConsoleMsg(Userindex, "Has click sobre un personaje antes!", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim midiID As Byte
        Dim mapa As Integer
        
        midiID = .incomingData.ReadByte
        mapa = .incomingData.ReadInteger
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, 50, 50) Then
                mapa = .Pos.map
            End If
        End If
    End With
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim waveID As Byte
        Dim mapa As Integer
        Dim X As Integer
        Dim Y As Integer
        
        waveID = .incomingData.ReadByte()
        mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadInteger()
        Y = .incomingData.ReadInteger()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
        'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, X, Y) Then
                mapa = .Pos.map
                X = .Pos.X
                Y = .Pos.Y
            End If
            
            'Ponemos el pedido por el GM
            Call SendToAreaByPosVisible(mapa, X, Y, PrepareMessagePlayWave(waveID, X, Y))
        End If
    End With
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("ARMADA REAL> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "CitizenMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCitizenMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "CriminalMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCriminalMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(Userindex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MargenY + 1 To .Pos.Y + MargenY - 1
            For X = .Pos.X - MargenX + 1 To .Pos.X + MargenX - 1
                If X > 0 And Y > 0 And X <= MapInfo(.Pos.map).Width And Y <= MapInfo(.Pos.map).Height Then
                    If MapData(.Pos.map).Tile(X, Y).ObjInfo.ObjIndex > 0 Then
                        If ItemNoEsDeMapa(MapData(.Pos.map).Tile(X, Y).ObjInfo.ObjIndex) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.map, X, Y)
                        End If
                    End If
                End If
            Next X
        Next Y
        
        Call LogGM(UserList(Userindex).Name, "/MASSDEST")
    End With
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False, True)
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim tObj As Integer
        Dim lista As String
        Dim X As Long
        Dim Y As Long
        
        For X = 5 To 95
            For Y = 5 To 95
                tObj = MapData(.Pos.map).Tile(X, Y).ObjInfo.ObjIndex
                If tObj > 0 Then
                    If ObjData(tObj).OBJType <> eOBJType.otArboles Then
                        Call WriteConsoleMsg(Userindex, "(" & X & "," & Y & ") " & ObjData(tObj).Name, FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Next Y
        Next X
    End With
End Sub

''
' Handles the "MakeDumb" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumb(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                Call FlushBuffer(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SecurityIp.DumpTables
    End With
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                If PersonajeExiste(UserName) Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline, Echando de los consejos", FontTypeNames.FONTTYPE_INFO)

                    modMySQL.Execute ("UPDATE pjs SET PerteneceConsejo=0, PerteneceCaos=0 WHERE Nombre='" & UserName & "'")
                Else
                    Call WriteConsoleMsg(Userindex, "No se encuentra el charfile " & UserName, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
                        Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del Concilio de las Sombras", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim tTrigger As Byte
        Dim tLog As String
        
        tTrigger = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If tTrigger >= 0 Then
            MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.X & "," & .Pos.Y
            
            Call LogGM(.Name, tLog)
            Call WriteConsoleMsg(Userindex, tLog, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'
'***************************************************
    Dim tTrigger As Byte
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        tTrigger = MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Trigger
        
        Call LogGM(.Name, "Miro el trigger en " & .Pos.map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(Userindex, _
            "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.X & ", " & .Pos.Y _
            , FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim lista As String
        Dim LoopC As Long
        
        Call LogGM(.Name, "/BANIPLIST")
        
        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(Userindex, lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call BanIpGuardar
        Call BanIpCargar
    End With
End Sub

''
' Handles the "GuildBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName As String
        Dim cantMembers As Integer
        Dim LoopC As Long
        Dim member As String
        Dim Count As Byte
        Dim tIndex As Integer
        Dim tFile As String
        Dim idPj As Long
        
        GuildName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(Userindex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " banned al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_FIGHT))
                
                'baneamos a los miembros
                Call LogGM(.Name, "BANCLAN a " & UCase$(GuildName))
                
                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                
                For LoopC = 1 To cantMembers
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    'member es la victima
                    Call Ban(member, "Administracion del servidor", "Clan Banned")
                    
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))
                    
                    tIndex = NameIndex(member)
                    If tIndex > 0 Then
                        'esta online
                        UserList(tIndex).flags.Ban = 1
                        Call CloseSocket(tIndex)
                    End If
                                  
                    idPj = val(GetByCampo("SELECT Id FROM pjs WHERE Nombre=" & Comillas(member), "Id"))
                    'ponemos el flag de ban a 1
                    modMySQL.Execute ("UPDATE pjs SET Penas=Penas+1, Ban=1 WHERE Id=" & idPj)
                    'ponemos la pena
                    modMySQL.Execute ("INSERT INTO penas (IdPj, Razon, Fecha, IdGM, Tiempo) VALUES (" & idPj & "," & Comillas(.Name & ": BAN AL CLAN: " & GuildName) & ",NOW()," & .MySQLId & ",0)")
                Next LoopC
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/12/08
'Agregado un CopyBuffer porque se producia un bucle
'inifito al intentar banear una ip ya baneada. (NicoNZ)
'***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim bannedIP As String
        Dim tUser As Integer
        Dim Reason As String
        Dim i As Long
        
        ' Is it by ip??
        If buffer.ReadBoolean() Then
            bannedIP = buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte()
        Else
            tUser = NameIndex(buffer.ReadASCIIString())
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
            Else
                bannedIP = UserList(tUser).ip
            End If
        End If
        
        Reason = buffer.ReadASCIIString()
        
        If LenB(bannedIP) > 0 Then
            If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                Call LogGM(.Name, "/BanIP " & bannedIP & " por " & Reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(Userindex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                    Call .incomingData.CopyBuffer(buffer) ' Agregado porque sino no se sacaba del
                                                          ' buffer y se hacia un bucle infinito. (NicoNZ) 05/12/2008
                    Exit Sub
                End If
                
                Call BanIpAgrega(bannedIP)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " baneó la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))
                
                'Find every player with that ip and ban him!
                For i = 1 To LastUser
                    If UserList(i).ConnIDValida Then
                        If UserList(i).ip = bannedIP Then
                            Call BanCharacter(Userindex, UserList(i).Name, "IP POR " & Reason)
                        End If
                    End If
                Next i
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim bannedIP As String
        
        bannedIP = .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(Userindex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Cant As Integer
        Dim ObjS As String
        Dim ObjIndex As Integer
        Dim Listado As String
        Dim found As Integer
        Cant = .incomingData.ReadInteger()
        
        ObjS = .incomingData.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        
        
        If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y - 1).TileExit.map > 0 Then _
            Exit Sub
        Dim i As Integer
        ObjIndex = val(ObjS)
        If ObjIndex = 0 Then
            For i = 1 To NumObjDatas
                If InStr(1, ObjData(i).Name, ObjS, vbTextCompare) > 0 Then
                    Listado = Listado & i & " - " & ObjData(i).Name & vbCrLf
                    found = found + 1
                    ObjIndex = i
                End If
            Next i
        
        
            If found = 0 Then
                Call WriteConsoleMsg(Userindex, "No se han encontrado objetos que contengan '" & ObjS & "'", FontTypeNames.FONTTYPE_GUILD)
                Exit Sub
            ElseIf found > 60 Then
                Call WriteConsoleMsg(Userindex, "Se han encontrado " & found & " objetos que contienen '" & ObjS & "', por favor acote el parámetro de búsqueda." & vbCrLf & Listado, FontTypeNames.FONTTYPE_GUILD)
                Exit Sub
            ElseIf found > 1 Then
                Call WriteConsoleMsg(Userindex, "Se han encontrado " & found & " objetos que contienen '" & ObjS & "':" & vbCrLf & Listado, FontTypeNames.FONTTYPE_GUILD)
                Exit Sub
            End If
        End If
        
        If ObjIndex < 1 Or ObjIndex > NumObjDatas Then _
            Exit Sub
        
        'Is the object not null?
        If LenB(ObjData(ObjIndex).Name) = 0 Then Exit Sub
        
        Call LogGM(.Name, "/CI: " & ObjIndex & " - " & ObjData(ObjIndex).Name & " - " & Cant)
        
        Dim Objeto As Obj
        Call WriteConsoleMsg(Userindex, "ATENCION: FUERON CREADOS ***" & Cant & "*** ITEMS!, TIRE Y /DEST LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
        
        Objeto.Amount = Cant
        Objeto.ObjIndex = ObjIndex
        Call MakeObj(Objeto, .Pos.map, .Pos.X, .Pos.Y - 1)
    End With
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub
        
        Call LogGM(.Name, "/DEST")
        
        If ObjData(MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then
            Call WriteConsoleMsg(Userindex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call EraseObj(10000, .Pos.map, .Pos.X, .Pos.Y)
    End With
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.Name, "ECHO DEL CAOS A: " & UserName)
    
            If tUser > 0 Then
                UserList(tUser).Faccion.FuerzasCaos = 0
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tUser)
            Else
                If PersonajeExiste(UserName) Then
                    modMySQL.Execute ("UPDATE pjs SET EjercitoCaos=0, Reenlistadas=200, Extra=" & Comillas("Expulsado por " & .Name) & " WHERE Nombre=" & Comillas(UserName))
                    Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, UserName & " inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.Name, "ECHO DE LA REAL A: " & UserName)
            
            If tUser > 0 Then
                UserList(tUser).Faccion.ArmadaReal = 0
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tUser)
            Else
                If PersonajeExiste(UserName) Then
                    modMySQL.Execute ("UPDATE pjs SET EjercitoReal=0, Reenlistadas=200, Extra=" & Comillas("Expulsado por " & .Name) & " WHERE Nombre=" & Comillas(UserName))
                    Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, UserName & " inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim waveID As Byte
        waveID = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
    End With
End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 1/05/07
'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
'***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim punishment As Byte
        Dim NewText As String
        
        UserName = buffer.ReadASCIIString()
        punishment = buffer.ReadByte
        NewText = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(Userindex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                If PersonajeExiste(UserName) Then
                    Call LogGM(.Name, " borro la pena: " & punishment & "-" & _
                      GetByCampo("SELECT Razon FROM penas WHERE Numero=" & val(punishment) & " AND Usuario=" & Comillas(UserName), "Razon") _
                      & " de " & UserName & " y la cambió por: " & NewText)
                    
                    modMySQL.Execute ("UPDATE Penas SET Razon=" & Comillas(.Name & ": <" & NewText & ">") & ", Fecha=" & FechaHora & " WHERE Numero=" & val(punishment) & " AND Usuario=" & Comillas(UserName))
                    
                    Call WriteConsoleMsg(Userindex, "Pena Modificada.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call LogGM(.Name, "/BLOQ")
        
        If MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Blocked = 0
        End If
        
        Call Bloquear(True, .Pos.map, .Pos.X, .Pos.Y, MapData(.Pos.map).Tile(.Pos.X, .Pos.Y).Blocked)
    End With
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.Name, "/MATA " & Npclist(.flags.TargetNPC).Name)
    End With
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MargenY To .Pos.Y + MargenY
            For X = .Pos.X - MargenX To .Pos.X + MargenX
                If X > 0 And Y > 0 And X <= MapInfo(.Pos.map).Width And Y <= MapInfo(.Pos.map).Height Then
                    If MapData(.Pos.map).Tile(X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.map).Tile(X, Y).NpcIndex)
                End If
            Next X
        Next Y
        Call LogGM(.Name, "/MASSKILL")
    End With
End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal Userindex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim lista As String
        Dim LoopC As Byte
        Dim priv As Integer
        Dim validCheck As Boolean
        
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            'Handle special chars
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")
            End If
            
            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            End If
            
            If validCheck Then
                Call LogGM(.Name, "/LASTIP " & UserName)
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conectó son:"
                    For LoopC = 1 To 5
                        lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
                    Next LoopC
                    Call WriteConsoleMsg(Userindex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(Userindex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Change the user`s chat color
'***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim color As Long
        
        color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = color
        End If
    End With
End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Ignore the user
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If
    End With
End Sub

''
' Handles the "CheckSlot" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal Userindex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 09/09/2008 (NicoNZ)
'Check one Users Slot in Particular from Inventory
'***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Slot As Byte
        Dim tIndex As Integer
        
        UserName = buffer.ReadASCIIString() 'Que UserName?
        Slot = buffer.ReadByte() 'Que Slot?
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Dios) Then
            tIndex = NameIndex(UserName)  'Que user index?
            
            Call LogGM(.Name, .Name & " Checkeo el slot " & Slot & " de " & UserName)
               
            If tIndex > 0 Then
                If Slot > 0 And Slot <= MAX_INVENTORY_SLOTS Then
                    If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
                        Call WriteConsoleMsg(Userindex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).Name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(Userindex, "No hay Objeto en slot seleccionado", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "Slot Inválido.", FontTypeNames.FONTTYPE_TALK)
                End If
            Else
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handles the "ResetAutoUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleResetAutoUpdate(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reset the AutoUpdate
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub
        
        Call WriteConsoleMsg(Userindex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Restart" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleRestart(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Restart the game
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub
        
        'time and Time BUG!
        Call LogGM(.Name, .Name & " reinicio el mundo")
        
        Call ReiniciarServidor(True)
    End With
End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the objects
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado a los objetos.")
        
        Call LoadOBJData
    End With
End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the spells
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los hechizos.")
        
        Call CargarHechizos
    End With
End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s INI
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los INITs.")
        
        Call LoadSini
    End With
End Sub

''
' Handle the "ReloadNPCs" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s NPC
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
         
        Call LogGM(.Name, .Name & " ha recargado los NPCs.")
    
        Call CargaNpcsDat
    
        Call WriteConsoleMsg(Userindex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Kick all the chars that are online
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados
    End With
End Sub


''
' Handle the "ShowServerForm" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Show the server form
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click
    End With
End Sub

''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Clean the SOS
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha borrado los SOS")
        
        Call Ayuda.Reset
    End With
End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Save the characters
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado todos los chars")
        
        Call mdParty.ActualizaExperiencias
        Call GuardarUsuarios
    End With
End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Saves the map
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.map))
        
        Call GrabarMapa(.Pos.map, WorldBkPath & "\Mapa" & .Pos.map & ".bak")
        
        Call WriteConsoleMsg(Userindex, "Mapa Guardado", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Allows admins to read guild messages
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        
        guild = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(Userindex, guild)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha hecho un backup")
        
        Call ES.HacerBackUp 'Sino lo confunde con la id del paquete
    End With
End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/26/06
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Activate or desactivate the Centinel
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        centinelaActivado = Not centinelaActivado
        
        With Centinela
            .RevisandoUserIndex = 0
            .clave = 0
            .TiempoRestante = 0
        End With
    
        If CentinelaNPCIndex Then
            Call QuitarNPC(CentinelaNPCIndex)
            CentinelaNPCIndex = 0
        End If
        
        If centinelaActivado Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido activado.", FontTypeNames.FONTTYPE_SERVER))
        Else
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido desactivado.", FontTypeNames.FONTTYPE_SERVER))
        End If
    End With
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user name
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the userName and newUser Packets
        Dim UserName As String
        Dim newName As String
        Dim changeNameUI As Integer
        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        newName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(Userindex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(Userindex, "El Pj esta online, debe salir para el cambio", FontTypeNames.FONTTYPE_WARNING)
                Else
                    If Not FileExist(CharPath & UserName & ".chr") Then
                        Call WriteConsoleMsg(Userindex, "El pj " & UserName & " es inexistente ", FontTypeNames.FONTTYPE_INFO)
                    Else
                        GuildIndex = val(GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX"))
                        
                        If GuildIndex > 0 Then
                            Call WriteConsoleMsg(Userindex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If Not FileExist(CharPath & newName & ".chr") Then
                                Call FileCopy(CharPath & UserName & ".chr", CharPath & UCase$(newName) & ".chr")
                                
                                Call WriteConsoleMsg(Userindex, "Transferencia exitosa", FontTypeNames.FONTTYPE_INFO)
                                
                                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                                
                                Dim cantPenas As Byte
                                
                                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.Name) & ": BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & time)
                                
                                Call LogGM(.Name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(Userindex, "El nick solicitado ya existe", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterMail(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim newMail As String
        
        UserName = buffer.ReadASCIIString()
        newMail = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
                Call WriteConsoleMsg(Userindex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(Userindex, "No existe el charfile " & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteVar(CharPath & UserName & ".chr", "CONTACTO", "Email", newMail)
                    Call WriteConsoleMsg(Userindex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)
                End If
                
                Call LogGM(.Name, "Le ha cambiado el mail a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "AlterPassword" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim copyFrom As String
        Dim password As String
        
        UserName = Replace(buffer.ReadASCIIString(), "+", " ")
        copyFrom = Replace(buffer.ReadASCIIString(), "+", " ")
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha alterado la contraseña de " & UserName)
            
            If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
                Call WriteConsoleMsg(Userindex, "usar /APASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not FileExist(CharPath & UserName & ".chr") Or Not FileExist(CharPath & copyFrom & ".chr") Then
                    Call WriteConsoleMsg(Userindex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
                Else
                    password = GetVar(CharPath & copyFrom & ".chr", "INIT", "Password")
                    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Password", password)
                    
                    Call WriteConsoleMsg(Userindex, "Password de " & UserName & " ha cambiado por la de " & copyFrom, FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False, UserList(Userindex).zona)
        
        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumoneo a " & Npclist(NpcIndex).Name & " en mapa " & .Pos.map)
        End If
    End With
End Sub


''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True, UserList(Userindex).zona)
        
        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumoneo con respawn " & Npclist(NpcIndex).Name & " en mapa " & .Pos.map)
        End If
    End With
End Sub

''
' Handle the "ImperialArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleImperialArmour(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim index As Byte
        Dim ObjIndex As Integer
        
        index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Select Case index
            Case 1
                ArmaduraImperial1 = ObjIndex
            
            Case 2
                ArmaduraImperial2 = ObjIndex
            
            Case 3
                ArmaduraImperial3 = ObjIndex
            
            Case 4
                TunicaMagoImperial = ObjIndex
        End Select
    End With
End Sub

''
' Handle the "ChaosArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChaosArmour(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim index As Byte
        Dim ObjIndex As Integer
        
        index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Select Case index
            Case 1
                ArmaduraCaos1 = ObjIndex
            
            Case 2
                ArmaduraCaos2 = ObjIndex
            
            Case 3
                ArmaduraCaos3 = ObjIndex
            
            Case 4
                TunicaMagoCaos = ObjIndex
        End Select
    End With
End Sub

''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/12/07
'
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1
        End If
        
        'Tell the client that we are navigating.
        Call WriteNavigateToggle(Userindex)
    End With
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(Userindex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
        Else
            Call WriteConsoleMsg(Userindex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
        End If
    End With
End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'Turns off the server
'***************************************************
    Dim handle As Integer
    
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "/APAGAR")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
        
        'Log
        handle = FreeFile
        Open LogPath & "\Main.log" For Append Shared As #handle
        
        Print #handle, Date & " " & time & " server apagado por " & .Name & ". "
        
        Close #handle
        
        Unload frmMain
    End With
End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/CONDEN " & UserName)
            
            tUser = NameIndex(UserName)
            If tUser > 0 Then _
                Call VolverCriminal(tUser)
        End If
                
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/09/09
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim Char As String
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJAR " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call ResetFacciones(tUser)
            Else
                Char = CharPath & UserName & ".chr"
                
                Call WriteVar(Char, "FACCIONES", "EjercitoReal", 0)
                Call WriteVar(Char, "FACCIONES", "CiudMatados", 0)
                Call WriteVar(Char, "FACCIONES", "CrimMatados", 0)
                Call WriteVar(Char, "FACCIONES", "EjercitoCaos", 0)
                Call WriteVar(Char, "FACCIONES", "FechaIngreso", "No ingresó a ninguna Facción")
                Call WriteVar(Char, "FACCIONES", "rArCaos", 0)
                Call WriteVar(Char, "FACCIONES", "rArReal", 0)
                Call WriteVar(Char, "FACCIONES", "rExCaos", 0)
                Call WriteVar(Char, "FACCIONES", "rExReal", 0)
                Call WriteVar(Char, "FACCIONES", "recCaos", 0)
                Call WriteVar(Char, "FACCIONES", "recReal", 0)
                Call WriteVar(Char, "FACCIONES", "Reenlistadas", 0)
                Call WriteVar(Char, "FACCIONES", "NivelIngreso", 0)
                Call WriteVar(Char, "FACCIONES", "MatadosIngreso", 0)
                Call WriteVar(Char, "FACCIONES", "NextRecompensa", 0)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJARCLAN " & UserName)
            
            GuildIndex = modGuilds.m_EcharMiembroDeClan(Userindex, UserName)
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(Userindex, "No pertenece a ningún clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "RequestCharMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Request user mail
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        Dim mail As String
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If FileExist(CharPath & UserName & ".chr") Then
                mail = GetVar(CharPath & UserName & ".chr", "CONTACTO", "email")
                
                Call WriteConsoleMsg(Userindex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/29/06
'Send a message to all the users
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        message = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Mensaje de sistema:" & message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(message))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "SetMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetMOTD(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 03/31/07
'Set the MOTD
'Modified by: Juan Martín Sotuyo Dodero (Maraxus)
'   - Fixed a bug that prevented from properly setting the new number of lines.
'   - Fixed a bug that caused the player to be kicked.
'***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim newMOTD As String
        Dim auxiliaryString() As String
        Dim LoopC As Long
        
        newMOTD = buffer.ReadASCIIString()
        auxiliaryString = Split(newMOTD, vbCrLf)
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha fijado un nuevo MOTD")
            
            MaxLines = UBound(auxiliaryString()) + 1
            
            ReDim MOTD(1 To MaxLines)
            
            Call WriteVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines", CStr(MaxLines))
            
            For LoopC = 1 To MaxLines
                Call WriteVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                
                MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
            Next LoopC
            
            Call WriteConsoleMsg(Userindex, "Se ha cambiado el MOTD con exito", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

''
' Handle the "ChangeMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín sotuyo Dodero (Maraxus)
'Last Modification: 12/29/06
'Change the MOTD
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Exit Sub
        End If
        
        Dim auxiliaryString As String
        Dim LoopC As Long
        
        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
        Next LoopC
        
        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = left$(auxiliaryString, Len(auxiliaryString) - 2)
            End If
        End If
        
        Call WriteShowMOTDEditionForm(Userindex, auxiliaryString)
    End With
End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal Userindex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(Userindex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Call WritePong(Userindex)
    End With
End Sub

''
' Handles the "RequestPartyForm" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyForm(ByVal Userindex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        If .PartyIndex > 0 Then
            Call WriteShowPartyForm(Userindex)
            
        Else
            Call WriteConsoleMsg(Userindex, "No perteneces a ningún grupo!", FontTypeNames.FONTTYPE_INFOBOLD)
        End If
    End With
End Sub
''
' Handle the "SetIniVar" message
'
' @param userIndex The index of the user sending the message

Private Sub HandleRetosAbrir(ByVal Userindex As Integer)
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        'No le dejo abrir el cuadro si están en la sala de retos
        If UserList(Userindex).SalaIndex = 0 Then
            Call WriteRetosAbrir(Userindex)
        End If
    End With
End Sub

Private Sub HandleIntercambiarInv(ByVal Userindex As Integer)
    Dim Slot1 As Byte, Slot2 As Byte
    Dim Banco As Boolean
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Slot1 = .incomingData.ReadByte
        Slot2 = .incomingData.ReadByte
        Banco = .incomingData.ReadBoolean
        
        If Banco Then
            Call IntercambiarBanco(Userindex, Slot1, Slot2)
        Else
            Call IntercambiarInventario(Userindex, Slot1, Slot2)
        End If
    End With
End Sub

Private Sub HandleRetosDecide(ByVal Userindex As Integer)
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        Call modRetos.RetosDecide(Userindex, .incomingData.ReadBoolean)
    End With
End Sub

Public Sub HandleRetosCrear(ByVal Userindex As Integer)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 21/06/09
'Modify server.ini
'***************************************************
    If UserList(Userindex).incomingData.Length < 11 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim R As tReto
        Dim RI As Integer
        Dim i As Integer
        Dim Validado As Boolean
        
        
        'Obtengo los parámetros
        R.Active = True
        R.Vs = buffer.ReadByte()
        R.PorItems = buffer.ReadBoolean()
        R.Oro = buffer.ReadLong()
    
        R.Pj(1) = Userindex
        R.APj(1) = True

        For i = 2 To IIf(R.Vs = 1, 2, 4)
            R.Pj(i) = NameIndex(buffer.ReadASCIIString())
        Next i
        
        If R.Oro > MAX_APUESTA_RETO Then
            Call WriteRetosRespuesta(Userindex, 6)
        ElseIf R.Oro + VALOR_RETOS > .Stats.GLD Then
            Call WriteRetosRespuesta(Userindex, 2)
        ElseIf R.Oro < MIN_APUESTA_RETO Then
            Call WriteRetosRespuesta(Userindex, 9)
        ElseIf R.Vs = 1 Then 'Para retos 1vs1
            If R.Pj(2) = 0 Then
                Call WriteRetosRespuesta(Userindex, 3)
            ElseIf R.Pj(1) = R.Pj(2) Then
                Call WriteRetosRespuesta(Userindex, 8)
            ElseIf Zonas(.zona).Segura = 0 Or _
                   Zonas(UserList(R.Pj(2)).zona).Segura = 0 Then
                Call WriteRetosRespuesta(Userindex, 3)
            ElseIf .RetoIndex > 0 Or _
                   UserList(R.Pj(2)).RetoIndex > 0 Then
                Call WriteRetosRespuesta(Userindex, 5)
            ElseIf UserList(R.Pj(1)).Stats.ELV < 15 Or _
                   UserList(R.Pj(2)).Stats.ELV < 15 Then
                Call WriteRetosRespuesta(Userindex, 16)
            Else
                Validado = True
            End If
        Else 'Para retos 2vs2
            If R.Pj(2) = 0 Or R.Pj(3) = 0 Or R.Pj(4) = 0 Then
                Call WriteRetosRespuesta(Userindex, 3)
            ElseIf R.Pj(1) = R.Pj(2) Or R.Pj(1) = R.Pj(3) Or R.Pj(1) = R.Pj(4) Or _
                   R.Pj(2) = R.Pj(3) Or R.Pj(2) = R.Pj(4) Or R.Pj(3) = R.Pj(4) Then
                Call WriteRetosRespuesta(Userindex, 8)
            ElseIf Zonas(.zona).Segura = 0 Or _
                   Zonas(UserList(R.Pj(2)).zona).Segura = 0 Or _
                   Zonas(UserList(R.Pj(3)).zona).Segura = 0 Or _
                   Zonas(UserList(R.Pj(4)).zona).Segura = 0 Then
                Call WriteRetosRespuesta(Userindex, 3)
            ElseIf .RetoIndex > 0 Or _
                   UserList(R.Pj(2)).RetoIndex > 0 Or _
                   UserList(R.Pj(3)).RetoIndex > 0 Or _
                   UserList(R.Pj(4)).RetoIndex > 0 Then
                Call WriteRetosRespuesta(Userindex, 5)
            ElseIf UserList(R.Pj(1)).Stats.ELV < 15 Or _
                   UserList(R.Pj(2)).Stats.ELV < 15 Or _
                   UserList(R.Pj(3)).Stats.ELV < 15 Or _
                   UserList(R.Pj(4)).Stats.ELV < 15 Then
                Call WriteRetosRespuesta(Userindex, 16)
            Else
                Validado = True
            End If
        End If

        
        If Validado Then
            
            For i = 1 To LastRetoIndex
                If Retos(i).Active = False Then Exit For
            Next i
            
            If i <= MAX_SOLRETOS Then
                If i > LastRetoIndex Then LastRetoIndex = i
                
                RI = i
                Retos(RI) = R
            
                .RetoIndex = RI
                For i = 2 To IIf(R.Vs = 1, 2, 4)
                    Call WriteConsoleMsg(R.Pj(i), UserList(Userindex).Name & " te ha invitado a participar en un reto, ve a la sección retos o escribe /RETOS para ver más información.", FontTypeNames.FONTTYPE_RETOS)
                    UserList(R.Pj(i)).RetoIndex = RI
                Next i
            
                'Reto creado con exito
                Call WriteRetosRespuesta(Userindex, 1)
                Call WriteConsoleMsg(Userindex, "Se han enviado las invitaciones para participar del reto a los demás jugadores. Puedes ver quienes ya han aceptado el reto desde la ventana de retos.", FontTypeNames.FONTTYPE_RETOS)
            Else
                'Demasiadas solicitudes de retos creadas
                Call WriteRetosRespuesta(Userindex, 7)
            End If
    
    
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub


Public Sub HandleSetIniVar(ByVal Userindex As Integer)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 21/06/09
'Modify server.ini
'***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim sLlave As String
        Dim sClave As String
        Dim sValor As String
        
        'Obtengo los parámetros
        sLlave = buffer.ReadASCIIString()
        sClave = buffer.ReadASCIIString()
        sValor = buffer.ReadASCIIString()
    
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            Dim sTmp As String
            'Obtengo el valor según llave y clave
            sTmp = GetVar(IniPath & "Server.ini", sLlave, sClave)
            
            'Si obtengo un valor escribo en el server.ini
            If LenB(sTmp) Then
                Call WriteVar(IniPath & "Server.ini", sLlave, sClave, sValor)
                Call LogGM(.Name, "Modificó en server.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & " por " & sValor)
                Call WriteConsoleMsg(Userindex, "Modificó " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

Public Sub HandleAddGM(ByVal Userindex As Integer)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 21/06/09
'Modify server.ini
'***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler
    With UserList(Userindex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Nombre As String
        Dim Rango As Byte
        
        'Obtengo los parámetros
        Rango = buffer.ReadByte()
        Nombre = buffer.ReadASCIIString()
    
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            Dim NumGM As Integer
            'Obtengo el valor según llave y clave
            NumGM = val(GetVar(IniPath & "Server.ini", "INIT", IIf(Rango = 2, "SemiDioses", "Consejeros"))) + 1
            Call WriteVar(IniPath & "Server.ini", "INIT", IIf(Rango = 2, "SemiDioses", "Consejeros"), NumGM)
            
            Call WriteVar(IniPath & "Server.ini", IIf(Rango = 2, "SemiDioses", "Consejeros"), IIf(Rango = 2, "Semidios" & NumGM, "Consejero" & NumGM), Nombre)
            'Si obtengo un valor escribo en el server.ini
            Call LogGM(.Name, "Agregó al " & IIf(Rango = 2, "semidios", "consejero") & " " & Nombre)
            Call WriteConsoleMsg(Userindex, "Se ha agregado a " & Nombre & " como " & IIf(Rango = 2, "semidios", "consejero"), FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
    End With

Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.Raise error
End Sub

Private Sub HandleEquitar(ByVal Userindex As Integer)

With UserList(Userindex)
    Call .incomingData.ReadByte
    
        If .flags.Equitando = True Then
                Call QuitarMontura(Userindex)
            Exit Sub
        End If
        
        If .Counters.Saliendo = True Then
            Call WriteConsoleMsg(Userindex, "No puedes realizar esta acción saliendo del juego!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Chiquito = True Then
            Call WriteConsoleMsg(Userindex, "No llegas a subirte al animal debido a tu apariencia pequeña!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(Userindex, "Estás comerciando!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
'        If .flags.TargetNPC > 0 Then
'            If Npclist(.flags.TargetNPC).Numero = NpcCaballo Then
'                If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 2 Then
'                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del animal salvaje.", FontTypeNames.FONTTYPE_INFO)
'                    Exit Sub
'                Else
'                    Call Montar(UserIndex, .flags.TargetNPC)
'                End If
'            Else
'                Call Montar(UserIndex, 0)
'            End If
'        Else
'            Call Montar(UserIndex, 0)
'        End If

    Call Montar(Userindex, 0)
   
End With
End Sub


Private Sub HandleDejarMontura(ByVal Userindex As Integer)
    With UserList(Userindex)
        Call .incomingData.ReadByte
        
            If .flags.Comerciando Then
                Call WriteConsoleMsg(Userindex, "Estás comerciando!!.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
              If .flags.Equitando = True Then
                    Call QuitarMontura(Userindex)
                Exit Sub
            End If
            
            
            Call LiberarMontura(Userindex)
            
       
    End With
End Sub

Private Sub HandleCreateEfectoClient(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Integer
        Dim Y As Integer
        Dim Effect As Byte
                
        X = .incomingData.ReadInteger()
        Y = .incomingData.ReadInteger()
        Effect = .incomingData.ReadByte()
        
        Select Case Effect
            Case eEffects.Bala
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateEfecto(UserList(Userindex).Char.CharIndex, 0, 7, 1, 27, Effect, .Pos.X, .Pos.Y, X, Y))
        End Select
        
    End With
End Sub



Private Sub HandleCreateEfectoClientAction(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.Raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Integer
        Dim Y As Integer
        Dim Effect As Integer
        
        X = .incomingData.ReadInteger()
        Y = .incomingData.ReadInteger()
        Effect = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Or Not InMapBounds(.Pos.map, X, Y) Then Exit Sub
        
        Select Case Effect
            Case eEffects.Bala
        
                'If Not InRangoVision(UserIndex, x, Y) Then
                '    Call WritePosUpdate(UserIndex)
                '    Exit Sub
                'End If
                
                'If exiting, cancel
                Call CancelExit(Userindex)
                    
                'Check attack interval
'                If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub
'                'Check Magic interval
'                If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Sub
'                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(Userindex) Then Exit Sub
                
                Call LanzarProyectil(Userindex, X, Y)
                
            End Select
   
    End With
End Sub


Private Sub HandleAnclarEmbarcacion(ByVal Userindex As Integer)

With UserList(Userindex)
    Call .incomingData.ReadByte
        If .Counters.Saliendo = True Then
                Call WriteConsoleMsg(Userindex, "No puedes realizar esta acción saliendo del juego!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        If .flags.Nadando = False Then
            If .flags.Navegando = 0 And .flags.Embarcado = 0 Then
                Call WriteConsoleMsg(Userindex, "No te encuentras en ninguna embarcación!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not TieneObjetos(474, 1, Userindex) And Not TieneObjetos(475, 1, Userindex) And Not TieneObjetos(476, 1, Userindex) Then ' Barca / Galera / Galeon
            Call WriteConsoleMsg(Userindex, "Para saltar al agua necesitas una embarcación!!.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
            Call Anclar_Embarcacion(Userindex)
        Else
            Call DesAnclar_Embarcacion(Userindex)
        End If
End With
End Sub


Public Sub WriteMultiMessage(ByVal Userindex As Integer, _
                             ByVal MessageIndex As Integer, _
                             Optional ByVal Arg1 As Long, _
                             Optional ByVal Arg2 As Long, _
                             Optional ByVal Arg3 As Long, _
                             Optional ByVal StringArg1 As String)

    On Error GoTo Errhandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.MultiMessage)
        Call .WriteByte(MessageIndex)
        
        Select Case MessageIndex
                Case eMessages.LanzaHechizoA
                    Call .WriteInteger(Arg1) 'Daño
                    Call .WriteASCIIString(StringArg1) 'Victima
                    Exit Sub
                Case eMessages.TeLanzanHechizo
                    Call .WriteInteger(Arg1) 'Daño
                    Call .WriteASCIIString(StringArg1) 'Atacante
                    Exit Sub
                Case eMessages.LeHasRobadoMonedasDeOro
                    Call .WriteInteger(Arg1) 'ORO
                    Call .WriteASCIIString(StringArg1)
                    Exit Sub
                Case eMessages.HanIntentadoRobarte
                    Call .WriteASCIIString(StringArg1) 'Persona
                    Exit Sub
                Case eMessages.NoTieneOro
                    Call .WriteASCIIString(StringArg1) 'Persona
                    Exit Sub
                Case eMessages.UserNoTieneObjetos
                    Call .WriteASCIIString(StringArg1) 'Persona
                    Exit Sub
                Case eMessages.HasRobadoObjetos
                    Call .WriteInteger(Arg1) 'Cantidad Objeto
                    Call .WriteASCIIString(StringArg1) 'Obj
                    Exit Sub
                 Case eMessages.HasHurtadoObjetos
                    Call .WriteInteger(Arg1) 'Cantidad Objeto
                    Call .WriteASCIIString(StringArg1) 'Obj
                    Exit Sub
                Case eMessages.HasApunaladoA
                    Call .WriteInteger(Arg1) 'Daño
                    Call .WriteASCIIString(StringArg1) 'Persona
                    Exit Sub
                Case eMessages.TeHanApunalado
                    Call .WriteInteger(Arg1) 'Daño
                    Call .WriteASCIIString(StringArg1) 'Persona
                    Exit Sub
                Case eMessages.HasApunaladoCriatura
                    Call .WriteInteger(Arg1) 'Daño
                    Exit Sub
                Case eMessages.HasGolpeadoCriticamente
                    Call .WriteInteger(Arg1) 'Daño
                    Call .WriteASCIIString(StringArg1) 'Persona
                    Exit Sub
                Case eMessages.TeHanGolpeadoCriticamente
                    Call .WriteInteger(Arg1) 'Daño
                    Call .WriteASCIIString(StringArg1) 'Persona
                    Exit Sub
                Case eMessages.HasGolpeadoCriticamenteCriatura
                    Call .WriteInteger(Arg1) 'Daño
                    Exit Sub
                Case eMessages.HasRecuperadoMana
                    Call .WriteInteger(Arg1) 'Mana
                    Exit Sub
        End Select

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume

    End If

End Sub


''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.


''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Logged" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Logged)
    Call UserList(Userindex).outgoingData.WriteByte(Hour(time))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteChangeHour(ByVal Userindex As Integer, Optional ByVal Sound As Boolean = True)

On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ChangeHour)
    Call UserList(Userindex).outgoingData.WriteByte(Hour(time))
    Call UserList(Userindex).outgoingData.WriteBoolean(Sound)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Public Function PrepareChangeHour() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ErrorMsg" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChangeHour)
        Call .WriteByte(Hour(time))
        
        PrepareChangeHour = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareCounterMsg(ByVal Counter As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ErrorMsg" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CuentaRegresiva)
        Call .WriteInteger(Counter)
        
        PrepareCounterMsg = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PreparePicInRender(ByVal picType As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PicInRender)
        Call .WriteInteger(picType)
        
        PreparePicInRender = .ReadASCIIStringFixed(.Length)
    End With
End Function


''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal Userindex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Disconnect" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Disconnect)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BankEnd)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Call UserList(Userindex).outgoingData.WriteLong(UserList(Userindex).Stats.Banco)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
    Call UserList(Userindex).outgoingData.WriteASCIIString(UserList(Userindex).ComUsu.DestNick)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
''
' Writes the "CancelOfferItem" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Slot      The slot to cancel.

Public Sub WriteCancelOfferItem(ByVal Userindex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/03/2010
'
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.CancelOfferItem)
        Call .WriteByte(Slot)
    End With
    
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Public Sub WriteCommerceChat(ByVal Userindex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)
'***************************************************
'Author: ZaMa
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareCommerceConsoleMsg(Chat, FontIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowBlacksmithForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowCarpenterForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "NPCSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCSwing(ByVal Userindex As Integer, Optional ByVal NpcType As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCSwing" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.NPCSwing)
    Call UserList(Userindex).outgoingData.WriteByte(NpcType)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCKillUser(ByVal Userindex As Integer, Optional ByVal NpcType As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCKillUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        .WriteByte (ServerPacketID.NPCKillUser)
        .WriteByte (NpcType)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldUser(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldUser)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldOther(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldother)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserSwing(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserSwing" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserSwing)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


''
' Writes the "SafeModeOn" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOn(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SafeModeOn" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.SafeModeOn)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOff(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SafeModeOff" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.SafeModeOff)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ResuscitationSafeOn" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationSafeOn(ByVal Userindex As Integer)
'***************************************************
'Author: Rapsodius
'Last Modification: 10/10/07
'Writes the "ResuscitationSafeOn" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ResuscitationSafeOn)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ResuscitationSafeOff" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationSafeOff(ByVal Userindex As Integer)
'***************************************************
'Author: Rapsodius
'Last Modification: 10/10/07
'Writes the "ResuscitationSafeOff" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ResuscitationSafeOff)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "NobilityLost" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNobilityLost(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NobilityLost" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.NobilityLost)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCantUseWhileMeditating(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.CantUseWhileMeditating)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(Userindex).Stats.MinSta)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(Userindex).Stats.MinMAN)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(Userindex).Stats.MinHP)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
''
' Writes the "UpdateBankGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateBankGold(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UpdateBankGold" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateBankGold)
        Call .WriteLong(UserList(Userindex).Stats.Banco)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateGold" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(Userindex).Stats.GLD)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateExp" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(Userindex).Stats.Exp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal Userindex As Integer, ByVal map As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMap" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteByte(map)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteInteger(UserList(Userindex).Pos.X)
        Call .WriteInteger(UserList(Userindex).Pos.Y)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteTooltip(ByVal Userindex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Tipo As Byte, ByVal Mensaje As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Tooltip)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteByte(Tipo)
        Call .WriteASCIIString(Mensaje)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WritePicInRender(ByVal Userindex As Integer, ByVal picType As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.PicInRender)
        Call .WriteByte(picType)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteQuit(ByVal Userindex As Integer, Optional ByVal Cancel As Byte = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Quit)
        Call .WriteByte(UserList(Userindex).Counters.Salir)
        Call .WriteByte(Cancel)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


''
' Writes the "NPCHitUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCHitUser(ByVal Userindex As Integer, ByVal Target As PartesCuerpo, ByVal damage As Integer, Optional ByVal NpcType As Byte = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCHitUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.NPCHitUser)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
        Call .WriteByte(NpcType)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserHitNPC" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    damage The number of HP lost by the target creature.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHitNPC(ByVal Userindex As Integer, ByVal damage As Integer, ByVal CharIndex As Integer, Optional ByVal NpcType As Byte = 0, Optional ByVal HitArea As Byte = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserHitNPC" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserHitNPC)
        
        'It is a long to allow the "drake slayer" (matadracos) to kill the great red dragon of one blow.
        Call .WriteInteger(damage)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(HitArea)
        Call .WriteByte(NpcType)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteUserSpellNPC(ByVal Userindex As Integer, ByVal damage As Integer, ByVal CharIndex As Integer, Optional ByVal NpcType As Byte = 0, Optional ByVal HitArea As Byte = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserHitNPC" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserSpellNPC)
        
        'It is a long to allow the "drake slayer" (matadracos) to kill the great red dragon of one blow.
        Call .WriteInteger(damage)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(HitArea)
        Call .WriteByte(NpcType)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserAttackedSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex       User to which the message is intended.
' @param    attackerIndex   The user index of the user that attacked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserAttackedSwing(ByVal Userindex As Integer, ByVal attackerIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserAttackedSwing" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserAttackedSwing)
        Call .WriteInteger(UserList(attackerIndex).Char.CharIndex)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserHittedByUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackerChar Char index of the user hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHittedByUser(ByVal Userindex As Integer, ByVal Target As PartesCuerpo, ByVal attackerChar As Integer, ByVal damage As Integer, Optional ByVal HitArea As Byte = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserHittedByUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserHittedByUser)
        Call .WriteInteger(attackerChar)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
        Call .WriteByte(HitArea)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserHittedUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackedChar Char index of the user hitted.
' @param    damage The number of HP lost by the oponent hitted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHittedUser(ByVal Userindex As Integer, ByVal Target As PartesCuerpo, ByVal attackedChar As Integer, ByVal damage As Integer, Optional ByVal HitArea As Byte = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserHittedUser" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserHittedUser)
        Call .WriteInteger(attackedChar)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
        Call .WriteByte(HitArea)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal Userindex As Integer, ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatOverHead" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(Chat, CharIndex, color))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal Userindex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(Chat, FontIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowBarco(ByVal Userindex As Integer, ByVal Sentido As Byte, ByVal mPaso As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal TiempoPuerto As Long, ByRef Pasajeros() As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
Dim i As Integer
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowBarco)
        Call .WriteByte(Sentido)
        Call .WriteByte(mPaso)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteLong(TiempoPuerto)
        For i = 1 To 4
            If Pasajeros(i) > 0 Then
                Call .WriteInteger(UserList(Pasajeros(i)).Char.CharIndex)
                Call .WriteInteger(UserList(Pasajeros(i)).Char.Body)
                Call .WriteInteger(UserList(Pasajeros(i)).Char.Head)
                Call .WriteInteger(UserList(Pasajeros(i)).Char.CascoAnim)
            Else
                Call .WriteInteger(0)
                Call .WriteInteger(0)
                Call .WriteInteger(0)
                Call .WriteInteger(0)
            End If
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Public Sub WriteQuitarBarco(ByVal Userindex As Integer, ByVal Sentido As Byte)

Dim i As Integer
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.QuitarBarco)
        Call .WriteByte(Sentido)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal Userindex As Integer, ByVal Chat As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildChat" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(Chat))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal Userindex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteShowMessageScroll(ByVal Userindex As Integer, ByVal message As String, ByVal Head As Integer, ByVal Opciones As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)
        Call .WriteInteger(Head)
        Call .WriteByte(Opciones)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(Userindex)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(Userindex).Char.CharIndex)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal Userindex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal Criminal As Byte, _
                                ByVal privileges As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterCreate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(Body, Head, heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, _
                                                            helmet, Name, Criminal, privileges))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal Userindex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterRemove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterMove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteForceCharMove(ByVal Userindex, ByVal Direccion As eHeading)
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Writes the "ForceCharMove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal Userindex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterChange" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(Body, Head, heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal Userindex As Integer, ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectCreate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(GrhIndex, X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal Userindex As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectDelete" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal Userindex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Blocked As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockPosition" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteBoolean(Blocked)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal Userindex As Integer, ByVal wave As Byte, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal Userindex As Integer, ByRef guildList() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim Tmp As String
    Dim i As Long
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AreaChanged" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteInteger(UserList(Userindex).Pos.X)
        Call .WriteInteger(UserList(Userindex).Pos.Y)
        Call .WriteByte(UserList(Userindex).Char.heading)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PauseToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RainToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageRainToggle())
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateFX" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(Userindex).Stats.MaxHP)
        Call .WriteInteger(UserList(Userindex).Stats.MinHP)
        Call .WriteInteger(UserList(Userindex).Stats.MaxMAN)
        Call .WriteInteger(UserList(Userindex).Stats.MinMAN)
        Call .WriteInteger(UserList(Userindex).Stats.MaxSta)
        Call .WriteInteger(UserList(Userindex).Stats.MinSta)
        Call .WriteLong(UserList(Userindex).Stats.GLD)
        Call .WriteByte(UserList(Userindex).Stats.ELV)
        Call .WriteLong(UserList(Userindex).Stats.ELU)
        Call .WriteLong(UserList(Userindex).Stats.Exp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteFirstInfo(ByVal Userindex As Integer)

On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.FirstInfo)
        Call .WriteByte(UserList(Userindex).clase)
        Call .WriteByte(IIf(Criminal(Userindex), 1, 2))
        
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal Userindex As Integer, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.WorkRequestTarget)
        Call .WriteByte(Skill)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "StopWorking" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.

Public Sub WriteStopWorking(ByVal Userindex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'
'***************************************************
On Error GoTo Errhandler
    
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.StopWorking)
        
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowPartyForm(ByVal Userindex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "ShowPartyForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    Dim PI As Integer
    Dim members(PARTY_MAXMEMBERS) As Integer
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowPartyForm)
        
        PI = UserList(Userindex).PartyIndex
        Call .WriteByte(CByte(Parties(PI).EsPartyLeader(Userindex)))
        
        If PI > 0 Then
            Call Parties(PI).ObtenerMiembrosOnline(members())
            For i = 1 To PARTY_MAXMEMBERS
                If members(i) > 0 Then
                    Tmp = Tmp & UserList(members(i)).Name & " (" & Fix(Parties(PI).MiExperiencia(members(i))) & ")" & SEPARATOR
                End If
            Next i
        End If
        
        If LenB(Tmp) <> 0 Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
            
        Call .WriteASCIIString(Tmp)
        Call .WriteLong(Parties(PI).ObtenerExperienciaTotal)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteRetosAbrir(ByVal Userindex As Integer)

On Error GoTo Errhandler
    Dim i As Long
    Dim RI As Integer
    Dim Acepto As Boolean

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.RetosAbre)
        
        RI = UserList(Userindex).RetoIndex
        If RI = 0 Then
            Call .WriteBoolean(True)
        Else
            Call .WriteBoolean(False)
            Call .WriteByte(Retos(RI).Vs)
            Call .WriteBoolean(Retos(RI).PorItems)
            Call .WriteLong(Retos(RI).Oro)
            For i = 1 To IIf(Retos(RI).Vs = 1, 2, 4)
                Call .WriteASCIIString(UserList(Retos(RI).Pj(i)).Name)
                Call .WriteBoolean(Retos(RI).APj(i))
                If Retos(RI).Pj(i) = Userindex And Retos(RI).APj(i) Then
                    Acepto = True
                End If
            Next i
            Call .WriteBoolean(Acepto)
        End If
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Public Sub WriteRetosRespuesta(ByVal Userindex As Integer, ByVal Respuesta As Byte)

On Error GoTo Errhandler
    
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.RetosRespuesta)
    Call UserList(Userindex).outgoingData.WriteByte(Respuesta)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteBorrarPJ(ByVal Userindex As Integer, ByVal UserName As String, ByVal PJName As String, ByVal password As String)
On Error GoTo Errhandler

'¿Es el passwd valido?
If password <> GetByCampo("SELECT Password FROM cuentas WHERE Nombre=" & Comillas(UserName), "Password") Then
    Call WriteErrorMsg(Userindex, "La contraseña es incorrecta.")
    Call FlushBuffer(Userindex)
    'Call CloseSocket(UserIndex)
    Exit Sub
End If

Dim idCuenta As Integer
Dim idPj As Integer
idCuenta = val(GetByCampo("SELECT id FROM cuentas WHERE Nombre=" & Comillas(UserName), "id"))
idPj = val(GetByCampo("SELECT id FROM pjs WHERE Nombre=" & Comillas(PJName) & " AND IdAccount= " & idCuenta, "id"))
If idPj > 0 Then
    modMySQL.Execute ("DELETE FROM pjs WHERE Id=" & idPj)
    Call WriteErrorMsg(Userindex, "El personaje ha sido borrado con éxito.")
    Call FlushBuffer(Userindex)
    Call CloseSocket(Userindex)
Else
    Call WriteErrorMsg(Userindex, "El personaje no existe o no te pertenece.")
    Call FlushBuffer(Userindex)
    Call CloseSocket(Userindex)
End If

Exit Sub

Errhandler:

End Sub

Public Sub WriteOpenAccount(ByVal Userindex As Integer, ByVal Name As String, ByVal password As String)
On Error GoTo Errhandler
'¿Es el passwd valido?

If password <> GetByCampo("SELECT Password FROM cuentas WHERE Nombre=" & Comillas(Name), "Password") Then
    Call WriteErrorMsg(Userindex, "Password incorrecto.")
    Call FlushBuffer(Userindex)
    Call CloseSocket(Userindex)
    Exit Sub
End If

Dim rs As clsMySQLRecordSet
Dim i As Integer

Dim Cant As Integer
Dim Head As Integer
Dim Body As Integer

Cant = mySQL.SQLQuery("SELECT pjs.Nombre, pjs.Head, pjs.Body,pjs.PerteneceReal, pjs.PerteneceCaos, pjs.Arma, pjs.Escudo, pjs.Casco, pjs.Logged, pjs.Muerto, pjs.Rep_Promedio, pjs.Clase, pjs.Navegando, pjs.ELV, pjs.GLD, pjs.Banco FROM pjs LEFT JOIN cuentas ON pjs.IdAccount=cuentas.id WHERE cuentas.Nombre=" & Comillas(Name) & " and Password=" & Comillas(password), rs)

    With UserList(Userindex).outgoingData

        Call .WriteByte(ServerPacketID.OpenAccount)
        Call .WriteInteger(Cant)
        For i = 0 To Cant - 1
            
            Call .WriteASCIIString(rs("Nombre"))
            If val(rs("Muerto")) = 0 Then
                Head = rs("Head")
                Body = rs("Body")
            Else
                If rs("Rep_Promedio") < 0 Then
                    Body = iCuerpoMuertoCrimi
                    Head = iCabezaMuertoCrimi
                Else
                    Body = iCuerpoMuerto
                    Head = iCabezaMuerto
                End If
            End If
            If CByte(rs("Navegando")) = 1 Then
                    Call .WriteInteger(0)
            Else
                  Call .WriteInteger(Head)
            End If
            Call .WriteInteger(Body)
            Call .WriteInteger(rs("Arma"))
            Call .WriteInteger(rs("Escudo"))
            Call .WriteInteger(rs("Casco"))
            Call .WriteBoolean(NameIndex(rs("Nombre")) > 0) 'val(rs("Logged")) = 1)
            Call .WriteBoolean(rs("Rep_Promedio") < 0)
            Call .WriteByte(rs("ELV")) ' Level
            Call .WriteByte(rs("Clase")) ' Clase
            Call .WriteLong(CLng(rs("GLD")) + CLng(rs("Banco"))) 'GLD
            
            If EsAdmin(rs("Nombre")) Then
                .WriteByte (PlayerType.Admin)
            ElseIf EsDios(rs("Nombre")) Then
                .WriteByte (PlayerType.Dios)
            ElseIf EsSemiDios(rs("Nombre")) Then
                .WriteByte (PlayerType.SemiDios)
            ElseIf EsConsejero(rs("Nombre")) Then
                .WriteByte (PlayerType.Consejero)
            Else
                If rs("PerteneceReal") = 1 Then
                    .WriteByte (PlayerType.RoyalCouncil)
                ElseIf rs("PerteneceCaos") = 1 Then
                    .WriteByte (PlayerType.ChaosCouncil)
                Else
                    .WriteByte (PlayerType.User)
                End If
                
            End If
            
            rs.MoveNext
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal Userindex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(Slot)
        
        Dim ObjIndex As Integer
        Dim obData As ObjData
        
        ObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call .WriteInteger(ObjIndex)
        Call .WriteASCIIString(obData.Name)
        Call .WriteInteger(UserList(Userindex).Invent.Object(Slot).Amount)
        Call .WriteBoolean(UserList(Userindex).Invent.Object(Slot).Equipped)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MinDef)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteSingle(SalePrice(obData.Valor))
        Call .WriteByte(UserNoPuedeUsarItem(Userindex, ObjIndex))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal Userindex As Integer, ByVal Slot As Integer, ByVal ObjIndex As Integer, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(Slot)
        
        Dim obData As ObjData
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)
        End If
        
        Call .WriteASCIIString(obData.Name)
        Call .WriteInteger(Amount)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.def)
        Call .WriteLong(obData.Valor)
        Call .WriteByte(UserNoPuedeUsarItem(Userindex, ObjIndex))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal Userindex As Integer, ByVal Slot As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(Slot)
        Call .WriteInteger(UserList(Userindex).Stats.UserHechizos(Slot))
        
        If UserList(Userindex).Stats.UserHechizos(Slot) > 0 Then
            Call .WriteASCIIString(Hechizos(UserList(Userindex).Stats.UserHechizos(Slot)).Nombre)
        Else
            Call .WriteASCIIString("(None)")
        End If
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal Userindex As Integer, ByVal SoloFA As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Atributes" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.atributes)
        If SoloFA Then
            Call .WriteByte(1)
            Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.fuerza))
            Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.agilidad))
            Call .WriteByte(UserList(Userindex).Stats.UserAtributosBackUP(eAtributos.fuerza))
            Call .WriteByte(UserList(Userindex).Stats.UserAtributosBackUP(eAtributos.agilidad))
        Else
            Call .WriteByte(2)
            Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.fuerza))
            Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.agilidad))
            Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia))
            Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma))
            Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion))
        End If
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)
        
        For i = 1 To UBound(ArmasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= Round(UserList(Userindex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(Userindex).clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(Obj.GrhIndex)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
            Call .WriteInteger(Obj.Upgrade)
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithArmors(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(Userindex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(Userindex).clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(Obj.GrhIndex)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
            Call .WriteInteger(Obj.Upgrade)
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjects)
        
        For i = 1 To UBound(ObjCarpintero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(Userindex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(Userindex).clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(Obj.GrhIndex)
            Call .WriteInteger(Obj.Madera)
            Call .WriteInteger(Obj.MaderaElfica)
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
            Call .WriteInteger(Obj.Upgrade)
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RestOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.RestOK)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal Userindex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ErrorMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(message))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Blind" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Blind)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Dumb" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Dumb)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSignal(ByVal Userindex As Integer, ByVal ObjIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSignal" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteASCIIString(ObjData(ObjIndex).texto)
        Call .WriteInteger(ObjData(ObjIndex).GrhSecundario)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal Userindex As Integer, ByVal Slot As Byte, ByRef Obj As Obj, ByVal price As Single, Optional ByVal PuedeUsarItem As Byte = 1)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 06/13/08
'Last Modified by: Nicolas Ezequiel Bouhid (NicoNZ)
'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim ObjInfo As ObjData
    
    If Obj.ObjIndex >= LBound(ObjData()) And Obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(Obj.ObjIndex)
    End If
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(Slot)
        Call .WriteASCIIString(ObjInfo.Name)
        Call .WriteInteger(Obj.Amount)
        Call .WriteSingle(price)
        Call .WriteInteger(ObjInfo.GrhIndex)
        Call .WriteInteger(Obj.ObjIndex)
        Call .WriteByte(ObjInfo.OBJType)
        Call .WriteInteger(ObjInfo.MaxHIT)
        Call .WriteInteger(ObjInfo.MinHIT)
        Call .WriteInteger(ObjInfo.def)
        Call .WriteByte(PuedeUsarItem)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub



''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(Userindex).Stats.MaxAGU)
        Call .WriteByte(UserList(Userindex).Stats.MinAGU)
        Call .WriteByte(UserList(Userindex).Stats.MaxHam)
        Call .WriteByte(UserList(Userindex).Stats.MinHam)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "Fame" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteFame(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Fame" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Fame)
        
        Call .WriteLong(UserList(Userindex).Reputacion.AsesinoRep)
        Call .WriteLong(UserList(Userindex).Reputacion.BandidoRep)
        Call .WriteLong(UserList(Userindex).Reputacion.BurguesRep)
        Call .WriteLong(UserList(Userindex).Reputacion.LadronesRep)
        Call .WriteLong(UserList(Userindex).Reputacion.NobleRep)
        Call .WriteLong(UserList(Userindex).Reputacion.PlebeRep)
        Call .WriteLong(UserList(Userindex).Reputacion.Promedio)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MiniStats" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        
        Call .WriteLong(UserList(Userindex).Faccion.CiudadanosMatados)
        Call .WriteLong(UserList(Userindex).Faccion.CriminalesMatados)
        
'TODO : Este valor es calculable, no debería NI EXISTIR, ya sea en el servidor ni en el cliente!!!
        Call .WriteLong(UserList(Userindex).Stats.UsuariosMatados)
        
        Call .WriteInteger(UserList(Userindex).Stats.NPCsMuertos)
        
        Call .WriteByte(UserList(Userindex).clase)
        Call .WriteLong(UserList(Userindex).Counters.Pena)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal Userindex As Integer, ByVal skillPoints As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LevelUp" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteGotHome(ByVal Userindex As Integer, ByVal ok As Boolean)
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.GotHome)
        Call .WriteBoolean(ok)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Public Sub WriteGoHome(ByVal Userindex As Integer, ByVal Tiempo As Integer)
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.goHome)
        Call .WriteInteger(Tiempo)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub


''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetInvisible" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

Public Sub WriteSetOculto(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetInvisible" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetOculto(CharIndex, invisible))
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MeditateToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlindNoMore" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumbNoMore" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SendSkills" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.SendSkills)
        Call .WriteInteger(UserList(Userindex).Stats.SkillPts)
        For i = 1 To NUMSKILLS
            Call .WriteByte(UserList(Userindex).Stats.UserSkills(i))
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim str As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To Npclist(NpcIndex).NroCriaturas
            str = str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then _
            str = left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal Userindex As Integer, ByVal guildNews As String, ByRef enemies() As String, ByRef allies() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildNews" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.guildNews)
        
        Call .WriteASCIIString(guildNews)
        
        'Prepare enemies' list
        For i = LBound(enemies()) To UBound(enemies())
            Tmp = Tmp & enemies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Tmp = vbNullString
        'Prepare allies' list
        For i = LBound(allies()) To UBound(allies())
            Tmp = Tmp & allies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
Public Sub WriteGuildMemberInfo(ByVal Userindex As Integer, ByRef guildList() As String, ByRef MemberList() As String)
'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'Writes the "GuildMemberInfo" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.GuildMemberInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal Userindex As Integer, ByVal details As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OfferDetails" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal Userindex As Integer, ByRef guilds() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)
        
        ' Prepare guild's list
        For i = LBound(guilds) To UBound(guilds)
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal Userindex As Integer, ByRef guilds() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)
                
        ' Prepare guilds' list
        For i = LBound(guilds) To UBound(guilds)
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "CharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
' @param    reputation The requested char's reputation.
' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterInfo(ByVal Userindex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, _
                            ByVal gender As eGenero, ByVal level As Byte, ByVal gold As Long, ByVal bank As Long, ByVal reputation As Long, _
                            ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, _
                            ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterInfo" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)
        
        Call .WriteASCIIString(charName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)
        
        Call .WriteByte(level)
        Call .WriteLong(gold)
        Call .WriteLong(bank)
        Call .WriteLong(reputation)
        
        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)
        
        Call .WriteBoolean(RoyalArmy)
        Call .WriteBoolean(CaosLegion)
        
        Call .WriteLong(citicensKilled)
        Call .WriteLong(criminalsKilled)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @param    guildNews The guild's news.
' @param    joinRequests The list of chars which requested to join the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeaderInfo(ByVal Userindex As Integer, ByRef guildList() As String, ByRef MemberList() As String, _
                            ByVal guildNews As String, ByRef joinRequests() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString
        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Store guild news
        Call .WriteASCIIString(guildNews)
        
        ' Prepare the join request's list
        Tmp = vbNullString
        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "GuildDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildName The requested guild's name.
' @param    founder The requested guild's founder.
' @param    foundationDate The requested guild's foundation date.
' @param    leader The requested guild's current leader.
' @param    URL The requested guild's website.
' @param    memberCount The requested guild's member count.
' @param    electionsOpen True if the clan is electing it's new leader.
' @param    alignment The requested guild's alignment.
' @param    enemiesCount The requested guild's enemy count.
' @param    alliesCount The requested guild's ally count.
' @param    antifactionPoints The requested guild's number of antifaction acts commited.
' @param    codex The requested guild's codex.
' @param    guildDesc The requested guild's description.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDetails(ByVal Userindex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, _
                            ByVal leader As String, ByVal URL As String, ByVal memberCount As Integer, ByVal electionsOpen As Boolean, _
                            ByVal alignment As String, ByVal enemiesCount As Integer, ByVal AlliesCount As Integer, _
                            ByVal antifactionPoints As String, ByRef Codex() As String, ByVal guildDesc As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildDetails" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim temp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)
        
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        Call .WriteASCIIString(URL)
        
        Call .WriteInteger(memberCount)
        Call .WriteBoolean(electionsOpen)
        
        Call .WriteASCIIString(alignment)
        
        Call .WriteInteger(enemiesCount)
        Call .WriteInteger(AlliesCount)
        
        Call .WriteASCIIString(antifactionPoints)
        
        For i = LBound(Codex()) To UBound(Codex())
            temp = temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(temp) > 1 Then _
            temp = left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
        
        Call .WriteASCIIString(guildDesc)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/12/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Writes the "ParalizeOK" message to the given user's outgoing data buffer
'And updates user position
'***************************************************
On Error GoTo Errhandler
        UserList(Userindex).outgoingData.WriteByte (ServerPacketID.ParalizeOK)
        Call WritePosUpdate(Userindex)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal Userindex As Integer, ByVal details As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TradeOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.TradeOK)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BankOK)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal Userindex As Integer, ByVal OfferSlot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
'25/11/2009: ZaMa - Now sends the specific offer slot to be modified.
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de sólo Def
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)
        
        Call .WriteByte(OfferSlot)
        Call .WriteInteger(ObjIndex)
        Call .WriteLong(Amount)
        
        If ObjIndex > 0 Then
            Call .WriteInteger(ObjData(ObjIndex).GrhIndex)
            Call .WriteByte(ObjData(ObjIndex).OBJType)
            Call .WriteInteger(ObjData(ObjIndex).MaxHIT)
            Call .WriteInteger(ObjData(ObjIndex).MinHIT)
            Call .WriteInteger(ObjData(ObjIndex).def)
            Call .WriteLong(SalePrice(ObjIndex))
            Call .WriteASCIIString(ObjData(ObjIndex).Name)
        Else ' Borra el item
            Call .WriteInteger(0)
            Call .WriteByte(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteLong(0)
            Call .WriteASCIIString("")
        End If
    End With
Exit Sub


Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal Userindex As Integer, ByRef npcNames() As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & npcNames(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) <> 0 Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal Userindex As Integer, ByVal currentMOTD As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)
        
        Call .WriteASCIIString(currentMOTD)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal Userindex As Integer, ByRef userNamesList() As String, ByVal Cant As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06 NIGO:
'Writes the "UserNameList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To Cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Pong" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Pong)
Exit Sub

Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With UserList(Userindex).outgoingData
        If .Length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.Length)
        
        Call EnviarDatosASlot(Userindex, sndData)
    End With
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "SetInvisible" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)
        
        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageSetOculto(ByVal CharIndex As Integer, ByVal invisible As Boolean) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "SetInvisible" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetOculto)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)
        
        PrepareMessageSetOculto = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ChatOverHead" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(CharIndex)
        
        ' Write rgb channels and save one byte from long :D
        Call .WriteByte(color And &HFF)
        Call .WriteByte((color And &HFF00&) \ &H100&)
        Call .WriteByte((color And &HFF0000) \ &H10000)
        
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal Chat As String, ByVal FontIndex As FontTypeNames) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ConsoleMsg" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageUsersOnline() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UsersOnline)
        Call .WriteInteger(NumUsers)
        
        PrepareMessageUsersOnline = .ReadASCIIStringFixed(.Length)
    End With
End Function
''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareCommerceConsoleMsg(ByRef Chat As String, ByVal FontIndex As FontTypeNames) As String
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'Prepares the "CommerceConsoleMsg" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CommerceChat)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareCommerceConsoleMsg = .ReadASCIIStringFixed(.Length)
    End With
End Function
Public Function PrepareMessageCreateEfecto(ByVal UserCharIndex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal Loops As Integer, ByVal Wav As Integer, ByVal Efecto As Byte, ByVal X As Integer, ByVal Y As Integer, Optional Xd As Integer = 0, Optional Xy As Integer = 0) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateEfecto)
        Call .WriteInteger(UserCharIndex)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(Loops)
        Call .WriteByte(Wav)
        Call .WriteByte(Efecto)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteInteger(Xd)
        Call .WriteInteger(Xy)
        
        PrepareMessageCreateEfecto = .ReadASCIIStringFixed(.Length)
    End With
End Function
Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageCreateAreaFX(ByVal X As Integer, Y As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateAreaFX)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCreateAreaFX = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageAtaca(ByVal Userindex As Integer, ByVal Tipo As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ataca)
        Call .WriteInteger(UserList(Userindex).Char.CharIndex)
        Call .WriteByte(Tipo)
        
        PrepareMessageAtaca = .ReadASCIIStringFixed(.Length)
    End With
End Function
Public Function PrepareMessageNpcAtaca(ByVal NpcIndex As Integer, ByVal Tipo As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ataca)
        Call .WriteInteger(Npclist(NpcIndex).Char.CharIndex)
        Call .WriteByte(Tipo)
        
        PrepareMessageNpcAtaca = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Byte, ByVal X As Integer, ByVal Y As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteByte(wave)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal Chat As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.GuildChat)
        Call .WriteASCIIString(Chat)
        
        PrepareMessageGuildChat = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal Chat As String) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "ShowMessageBox" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Chat)
        
        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageShowMessageScroll(ByVal Chat As String, ByVal Head As Integer, ByVal Opciones) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "ShowMessageBox" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageScroll)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(Head)
        Call .WriteByte(Opciones)
        
        PrepareMessageShowMessageScroll = .ReadASCIIStringFixed(.Length)
    End With
End Function


''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "PauseToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRainToggle() As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "RainToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RainToggle)
        Call .WriteByte(LloviendoConTormenta)
        
        PrepareMessageRainToggle = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Integer, ByVal Y As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ObjectDelete" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        
        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Integer, ByVal Y As Integer, ByVal Blocked As Boolean) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "BlockPosition" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteBoolean(Blocked)
        
        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.Length)
    End With
    
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'prepares the "ObjectCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteInteger(GrhIndex)
        
        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterRemove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal Body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal Criminal As Byte, _
                                ByVal privileges As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(Body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(Name)
        Call .WriteByte(Criminal)
        Call .WriteByte(privileges)
        
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal Body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterChange" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(Body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        
        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Prepares the "ForceCharMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)
        
        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.Length)
    End With
End Function
Public Function PrepareMessageAgregarPasajero(ByVal Userindex As Integer, ByVal Sentido As Byte, ByVal Lugar As Byte) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.AgregarPasajero)
        Call .WriteByte(Sentido)
        Call .WriteInteger(UserList(Userindex).Char.CharIndex)
        Call .WriteByte(Lugar)
        Call .WriteInteger(UserList(Userindex).Char.Head)
        Call .WriteInteger(UserList(Userindex).Char.Body)
        Call .WriteInteger(UserList(Userindex).Char.CascoAnim)
        
        
        PrepareMessageAgregarPasajero = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageShowBarco(ByVal Sentido As Byte, ByVal mPaso As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal TiempoPuerto As Long, ByRef Pasajeros() As Integer) As String
Dim i As Integer
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowBarco)
        Call .WriteByte(Sentido)
        Call .WriteByte(mPaso)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteLong(TiempoPuerto)
        
        For i = 1 To 4
        
            If Pasajeros(i) > 0 Then
                Call .WriteInteger(UserList(Pasajeros(i)).Char.CharIndex)
                Call .WriteInteger(UserList(Pasajeros(i)).Char.Body)
                Call .WriteInteger(UserList(Pasajeros(i)).Char.Head)
                Call .WriteInteger(UserList(Pasajeros(i)).Char.CascoAnim)
            Else
                Call .WriteInteger(0)
                Call .WriteInteger(0)
                Call .WriteInteger(0)
                Call .WriteInteger(0)
            End If
        Next i
        
        
        PrepareMessageShowBarco = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageQuitarPasajero(ByVal Userindex As Integer, ByVal Sentido As Byte, ByVal Lugar As Byte) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.QuitarPasajero)
        Call .WriteByte(Sentido)
        Call .WriteByte(Lugar)
        
        PrepareMessageQuitarPasajero = .ReadASCIIStringFixed(.Length)
    End With
End Function
''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal Userindex As Integer, isCriminal As Boolean, Tag As String) As String
'***************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modification: 04/07/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'Prepares the "UpdateTagAndStatus" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        
        Call .WriteInteger(UserList(Userindex).Char.CharIndex)
        Call .WriteBoolean(isCriminal)
        Call .WriteASCIIString(Tag)
        
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal message As String) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ErrorMsg" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ErrorMsg)
        Call .WriteASCIIString(message)
        
        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.Length)
    End With
End Function


Public Function PrepareMessageEquitando(ByVal CharIndex As Integer, ByVal Equitando As Boolean) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "SetInvisible" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetEquitando)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(Equitando)
        
        PrepareMessageEquitando = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageCongelado(ByVal CharIndex As Integer, ByVal Congelado As Boolean) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "SetInvisible" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetCongelado)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(Congelado)
        
        PrepareMessageCongelado = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageChiquito(ByVal CharIndex As Integer, ByVal Chiquito As Boolean) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetChiquito)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(Chiquito)
        
        PrepareMessageChiquito = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessagePalabrasMagicas(ByVal SpellIndex As String, _
                                              ByVal CharIndex As Integer) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PalabrasMagicas)
        Call .WriteASCIIString(SpellIndex)
        Call .WriteInteger(CharIndex)
     
        PrepareMessagePalabrasMagicas = .ReadASCIIStringFixed(.Length)

    End With

End Function

Public Function PrepareMessageNadando(ByVal CharIndex As Integer, ByVal Nadando As Boolean) As String
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetNadando)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(Nadando)
        
        PrepareMessageNadando = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Sub WriteQuestDetails(ByVal Userindex As Integer, ByVal QuestIndex As Integer, Optional QuestSlot As Byte = 0)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Envía el paquete QuestDetails y la información correspondiente.
'Last modified: 30/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
 
On Error GoTo Errhandler
    With UserList(Userindex).outgoingData
        'ID del paquete
        Call .WriteByte(ServerPacketID.QuestDetails)
        
        'Se usa la variable QuestSlot para saber si enviamos la info de una quest ya empezada o la info de una quest que no se aceptó todavía (1 para el primer caso y 0 para el segundo)
        Call .WriteByte(IIf(QuestSlot, 1, 0))
        
        'Enviamos nombre, descripción y nivel requerido de la quest
        Call .WriteASCIIString(QuestList(QuestIndex).Nombre)
        Call .WriteASCIIString(QuestList(QuestIndex).desc)
        Call .WriteByte(QuestList(QuestIndex).RequiredLevel)
        
        'Enviamos la cantidad de npcs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredNPCs)
        If QuestList(QuestIndex).RequiredNPCs Then
            'Si hay npcs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredNPCs
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).Amount)
                Call .WriteASCIIString(GetVar(DatPath & "NPCs.dat", "NPC" & QuestList(QuestIndex).RequiredNPC(i).NpcIndex, "Name"))
                'Si es una quest ya empezada, entonces mandamos los NPCs que mató.
                If QuestSlot Then
                    Call .WriteInteger(UserList(Userindex).QuestStats.Quests(QuestSlot).NPCsKilled(i))
                End If
            Next i
        End If
        
        'Enviamos la cantidad de objs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredOBJs)
        If QuestList(QuestIndex).RequiredOBJs Then
            'Si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredOBJs
                Call .WriteInteger(QuestList(QuestIndex).RequiredOBJ(i).Amount)
                Call .WriteASCIIString(ObjData(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex).Name)
            Next i
        End If
    
        'Enviamos la recompensa de oro y experiencia.
        Call .WriteLong(QuestList(QuestIndex).RewardGLD)
        Call .WriteLong(QuestList(QuestIndex).RewardEXP)
        
        'Enviamos la cantidad de objs de recompensa
        Call .WriteByte(QuestList(QuestIndex).RewardOBJs)
        If QuestList(QuestIndex).RequiredOBJs Then
            'si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RewardOBJs
                Call .WriteInteger(QuestList(QuestIndex).RewardOBJ(i).Amount)
                Call .WriteASCIIString(ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).Name)
            Next i
        End If
    End With
Exit Sub
 
Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub
 
Public Sub WriteQuestListSend(ByVal Userindex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Envía el paquete QuestList y la información correspondiente.
'Last modified: 30/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
Dim tmpStr As String
Dim tmpByte As Byte
 
On Error GoTo Errhandler
 
    With UserList(Userindex)
        .outgoingData.WriteByte ServerPacketID.QuestListSend
    
        For i = 1 To MAXUSERQUESTS
            If .QuestStats.Quests(i).QuestIndex Then
                tmpByte = tmpByte + 1
                tmpStr = tmpStr & QuestList(.QuestStats.Quests(i).QuestIndex).Nombre & "-"
            End If
        Next i
        
        'Escribimos la cantidad de quests
        Call .outgoingData.WriteByte(tmpByte)
        
        'Escribimos la lista de quests (sacamos el último caracter)
        If tmpByte Then
            Call .outgoingData.WriteASCIIString(left$(tmpStr, Len(tmpStr) - 1))
        End If
    End With
Exit Sub
 
Errhandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
        Resume
    End If
End Sub

