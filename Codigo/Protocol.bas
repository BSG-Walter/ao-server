Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martin Sotuyo Dodero (Maraxus)
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
'The binary prtocol here used was designed by Juan Martin Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Martin Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

#If False Then

    Dim Map, X, Y, n, Mapa, race, helmet, weapon, shield, color, Value, errHandler, punishments, Length, obj, index As Variant

#End If

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private Enum ServerPacketID
    Logged = 1                  ' LOGGED
    RemoveDialogs = 2           ' QTDL
    RemoveCharDialog = 3        ' QDL
    NavigateToggle = 4          ' NAVEG
    Disconnect = 5              ' FINOK
    CommerceEnd = 6             ' FINCOMOK
    BankEnd = 7                 ' FINBANOK
    CommerceInit = 8            ' INITCOM
    BankInit = 9                ' INITBANCO
    UserCommerceInit = 10        ' INITCOMUSU
    UserCommerceEnd = 11         ' FINCOMUSUOK
    UserOfferConfirm = 12
    CommerceChat = 13
    UpdateSta = 14               ' ASS
    UpdateMana = 15             ' ASM
    UpdateHP = 16                ' ASH
    UpdateGold = 17              ' ASG
    UpdateBankGold = 18
    UpdateExp = 19               ' ASE
    ChangeMap = 20               ' CM
    PosUpdate = 21              ' PU
    ChatOverHead = 22            ' ||
    ConsoleMsg = 23              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat = 24               ' |+
    ShowMessageBox = 25          ' !!
    UserIndexInServer = 26       ' IU
    UserCharIndexInServer = 27   ' IP
    CharacterCreate = 28         ' CC
    CharacterRemove = 29         ' BP
    CharacterChangeNick = 30
    CharacterMove = 31           ' MP, +, * and _ '
    ForceCharMove = 32
    CharacterChange = 33         ' CP
    HeadingChange = 34
    ObjectCreate = 35            ' HO
    ObjectDelete = 36            ' BO
    BlockPosition = 37           ' BQ
    PlayMp3 = 38
    PlayMidi = 39                ' TM
    PlayWave = 40                ' TW
    guildList = 41               ' GL
    AreaChanged = 42             ' CA
    PauseToggle = 43             ' BKW
    RainToggle = 44              ' LLU
    CreateFX = 45                ' CFX
    UpdateUserStats = 46         ' EST
    ChangeInventorySlot = 47     ' CSI
    ChangeBankSlot = 48          ' SBO
    ChangeSpellSlot = 49         ' SHS
    Atributes = 50               ' ATR
    BlacksmithWeapons = 51       ' LAH
    BlacksmithArmors = 52        ' LAR
    InitCarpenting = 53          ' OBR
    RestOK = 54                  ' DOK
    errorMsg = 55                ' ERR
    Blind = 56                   ' CEGU
    Dumb = 57                    ' DUMB
    ShowSignal = 58              ' MCAR
    ChangeNPCInventorySlot = 59  ' NPCI
    UpdateHungerAndThirst = 60   ' EHYS
    Fame = 61                    ' FAMA
    MiniStats = 62               ' MEST
    LevelUp = 63                 ' SUNI
    AddForumMsg = 64             ' FMSG
    ShowForumForm = 65           ' MFOR
    SetInvisible = 66            ' NOVER
    DiceRoll = 67                ' DADOS
    MeditateToggle = 68          ' MEDOK
    BlindNoMore = 69             ' NSEGUE
    DumbNoMore = 70              ' NESTUP
    SendSkills = 71              ' SKILLS
    TrainerCreatureList = 72     ' LSTCRI
    guildNews = 73               ' GUILDNE
    OfferDetails = 74            ' PEACEDE & ALLIEDE
    AlianceProposalsList = 75    ' ALLIEPR
    PeaceProposalsList = 76      ' PEACEPR
    CharacterInfo = 77           ' CHRINFO
    GuildLeaderInfo = 78         ' LEADERI
    GuildMemberInfo = 79
    GuildDetails = 80            ' CLANDET
    ShowGuildFundationForm = 81  ' SHOWFUN
    ParalizeOK = 82              ' PARADOK
    ShowUserRequest = 83         ' PETICIO
    ChangeUserTradeSlot = 84     ' COMUSUINV
    SendNight = 85               ' NOC
    Pong = 86
    UpdateTagAndStatus = 87
    
    'GM =  messages
    SpawnList = 88               ' SPL
    ShowSOSForm = 89             ' MSOS
    ShowMOTDEditionForm = 90     ' ZMOTD
    ShowGMPanelForm = 91         ' ABPANEL
    UserNameList = 92            ' LISTUSU
    ShowDenounces = 93
    RecordList = 94
    RecordDetails = 95
    
    ShowGuildAlign = 96
    ShowPartyForm = 97
    UpdateStrenghtAndDexterity = 98
    UpdateStrenght = 99
    UpdateDexterity = 100
    AddSlots = 101
    MultiMessage = 102
    StopWorking = 103
    CancelOfferItem = 104
    PalabrasMagicas = 105
    PlayAttackAnim = 106
    FXtoMap = 107
    AccountLogged = 108 'CHOTS | Accounts
    SearchList = 109
    QuestDetails = 110
    QuestListSend = 111
    CreateDamage = 112           ' CDMG
    UserInEvent = 113
    RenderMsg = 114
    DeletedChar = 115
    EquitandoToggle = 116
    EnviarDatosServer = 117
    InitCraftman = 118
    EnviarListDeAmigos = 119
    SeeInProcess = 120
    ShowProcess = 121
    proyectil = 122
    PlayIsInChatMode = 123
End Enum

Private Enum ClientPacketID

    LoginExistingChar = 1     'OLOGIN
    ThrowDices = 2            'TIRDAD
    LoginNewChar = 3          'NLOGIN
    Talk = 4                  ';
    Yell = 5                  '-
    Whisper = 6                 '\
    Walk = 7                     'M
    RequestPositionUpdate = 8    'RPU
    Attack = 9                  'AT
    PickUp = 10                   'AG
    SafeToggle = 11              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle = 12
    RequestGuildLeaderInfo = 13   'GLINFO
    RequestAtributes = 14         'ATR
    RequestFame = 15              'FAMA
    RequestSkills = 16            'ESKI
    RequestMiniStats = 17         'FEST
    CommerceEnd = 18             'FINCOM
    UserCommerceEnd = 19         'FINCOMUSU
    UserCommerceConfirm = 20
    CommerceChat = 21
    BankEnd = 22                'FINBAN
    UserCommerceOk = 23           'COMUSUOK
    UserCommerceReject = 24       'COMUSUNO
    Drop = 25                   'TI
    CastSpell = 26                'LH
    LeftClick = 27               'LC
    DoubleClick = 28             'RC
    Work = 29                     'UK
    UseSpellMacro = 30           'UMH
    UseItem = 31              'USA
    CraftBlacksmith = 32          'CNS
    CraftCarpenter = 33           'CNC
    WorkLeftClick = 34           'WLC
    CreateNewGuild = 35           'CIG
    sadasdA = 36
    EquipItem = 37               'EQUI
    ChangeHeading = 38           'CHEA
    ModifySkills = 39             'SKSE
    Train = 40                   'ENTR
    CommerceBuy = 41              'COMP
    BankExtractItem = 42          'RETI
    CommerceSell = 43            'VEND
    BankDeposit = 44              'DEPO
    ForumPost = 45                'DEMSG
    MoveSpell = 46               'DESPHE
    MoveBank = 47
    ClanCodexUpdate = 48         'DESCOD
    UserCommerceOffer = 49        'OFRECER
    GuildAcceptPeace = 50         'ACEPPEAT
    GuildRejectAlliance = 51      'RECPALIA
    GuildRejectPeace = 52        'RECPPEAT
    GuildAcceptAlliance = 53      'ACEPALIA
    GuildOfferPeace = 54          'PEACEOFF
    GuildOfferAlliance = 55       'ALLIEOFF
    GuildAllianceDetails = 56     'ALLIEDET
    GuildPeaceDetails = 57        'PEACEDET
    GuildRequestJoinerInfo = 58   'ENVCOMEN
    GuildAlliancePropList = 59    'ENVALPRO
    GuildPeacePropList = 60       'ENVPROPP
    GuildDeclareWar = 61          'DECGUERR
    GuildNewWebsite = 62          'NEWWEBSI
    GuildAcceptNewMember = 63     'ACEPTARI
    GuildRejectNewMember = 64     'RECHAZAR
    GuildKickMember = 65         'ECHARCLA
    GuildUpdateNews = 66          'ACTGNEWS
    GuildMemberInfo = 67          '1HRINFO<
    GuildOpenElections = 68       'ABREELEC
    GuildRequestMembership = 69   'SOLICITUD
    GuildRequestDetails = 70      'CLANDETAILS
    Online = 71                  '/ONLINE
    Quit = 72                     '/SALIR
    GuildLeave = 73               '/SALIRCLAN
    RequestAccountState = 74      '/BALANCE
    PetStand = 75                 '/QUIETO
    PetFollow = 76                '/ACOMPANAR
    ReleasePet = 77              '/LIBERAR
    TrainList = 78                '/ENTRENAR
    Rest = 79                     '/DESCANSAR
    Meditate = 80                '/MEDITAR
    Resucitate = 81               '/RESUCITAR
    Heal = 82                     '/CURAR
    Help = 83                    '/AYUDA
    RequestStats = 84             '/EST
    CommerceStart = 85           '/COMERCIAR
    BankStart = 86               '/BOVEDA
    Enlist = 87                   '/ENLISTAR
    Information = 88            '/INFORMACION
    Reward = 89                   '/RECOMPENSA
    RequestMOTD = 90              '/MOTD
    UpTime = 91                   '/UPTIME
    PartyLeave = 92               '/SALIRPARTY
    PartyCreate = 93              '/CREARPARTY
    PartyJoin = 94                '/PARTY
    Inquiry = 95                  '/ENCUESTA ( with no params )
    GuildMessage = 96             '/CMSG
    PartyMessage = 97             '/PMSG
    GuildOnline = 98              '/ONLINECLAN
    PartyOnline = 99             '/ONLINEPARTY
    CouncilMessage = 100           '/BMSG
    RoleMasterRequest = 101     '/ROL
    GMRequest = 102              '/GM
    bugReport = 103              '/_BUG
    ChangeDescription = 104      '/DESC
    GuildVote = 105              '/VOTO
    punishments = 106           '/PENAS
    ChangePassword = 107         '/CONTRASENA
    Gamble = 108                '/APOSTAR
    InquiryVote = 109            '/ENCUESTA ( with parameters )
    LeaveFaction = 110          '/RETIRAR ( with no arguments )
    BankExtractGold = 111        '/RETIRAR ( with arguments )
    BankDepositGold = 112        '/DEPOSITAR
    Denounce = 113               '/DENUNCIAR
    GuildFundate = 114          '/FUNDARCLAN
    GuildFundation = 115
    PartyKick = 116              '/ECHARPARTY
    PartySetLeader = 117         '/PARTYLIDER
    PartyAcceptMember = 118      '/ACCEPTPARTY
    Ping = 119                  '/PING
    RequestPartyForm = 120
    ItemUpgrade = 121
    GMCommands = 122
    InitCrafting = 123
    Home = 124
    ShowGuildNews = 125
    ShareNpc = 126               '/COMPARTIR
    StopSharingNpc = 127
    Consultation = 128
    moveItem = 129
    LoginExistingAccount = 130  'CHOTS | Accounts
    LoginNewAccount = 131       'CHOTS | Accounts
    CentinelReport = 132
    Ecvc = 133
    Acvc = 134
    IrCvc = 135
    DragAndDropHechizos = 136
    Quest = 137                  '/QUEST
    QuestAccept = 138
    QuestListRequest = 139
    QuestDetailsRequest = 140
    QuestAbandon = 141
    CambiarContrasena = 142
    FightSend = 143
    FightAccept = 144
    CloseGuild = 145
    Discord = 146
    DeleteChar = 147
    ObtenerDatosServer = 148
    CraftsmanCreate = 149
    AddAmigos = 150
    DelAmigos = 151
    OnAmigos = 152
    MsgAmigos = 153
    LookProcess = 154
    SendProcessList = 155
    SendIfCharIsInChatMode = 156
End Enum

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 156

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
    FONTTYPE_CRIMINAL

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
    eo_Vida
    eo_Poss

End Enum

Public Reader  As Network.Reader
Public Writer As Network.Writer

Public Sub InitializePacketList()
    Set Writer = New Network.Writer
End Sub

''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal UserIndex As Integer, ByVal Message As Network.Reader) As Boolean
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/09/07
    '
    '***************************************************
    On Error GoTo HandleIncomingData_Err
    Set Reader = Message
    
    Dim actual_time As Long
    Dim performance_timer As Long
    actual_time = GetTickCount()
    performance_timer = actual_time
    
    With UserList(UserIndex)
        
        'Contamos cuantos paquetes recibimos.
        .Counters.PacketsTick = .Counters.PacketsTick + 1
        
        'Comento esto por ahora, por que cuando hago worldsave, envia mas paquetes en 40ms
        'y desconecta al pj, hay que reveer que hacer con esto y como solucionarlo.
        
        'Si recibis 10 paquetes en 40ms (intervalo del GameTimer), cierro la conexion.
        'If .Counters.PacketsTick > 10 Then
        '    Call CloseSocket(Userindex)
        '    Exit Function
        
        'End If
        
        'Se castea a long por que VB6 cuando usa SELECT CASE
        'Lo hace de manera mas efectiva https://www.gs-zone.org/temas/las-consecuencias-de-usar-byte-en-handleincomingdata.99245/
        Dim packetID As Long: packetID = CLng(Reader.ReadInt8)
        
        'Verifico si el paquete necesita que el user este logeado
        If Not (packetID = ClientPacketID.ThrowDices _
        Or packetID = ClientPacketID.LoginExistingChar _
        Or packetID = ClientPacketID.LoginNewChar _
        Or packetID = ClientPacketID.LoginNewAccount _
        Or packetID = ClientPacketID.LoginExistingAccount _
        Or packetID = ClientPacketID.DeleteChar _
        Or packetID = ClientPacketID.ObtenerDatosServer _
        Or packetID = ClientPacketID.CambiarContrasena) Then
        
        'Vierifico si el user esta logeado
        If Not .flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Function
            
            'El usuario ya logueo. Reseteamos el tiempo AFK si el ID es valido.
        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
            .Counters.IdleCount = 0
            
        End If
        
    ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
        .Counters.IdleCount = 0
        
        'Vierifico si el user esta logeado
        If .flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Function
            
        End If
        
    End If
    
    ' Ante cualquier paquete, pierde la proteccion de ser atacado.
    .flags.NoPuedeSerAtacado = False
    
End With

Select Case packetID
    
    Case ClientPacketID.SendIfCharIsInChatMode        '... Chat Mode
        Call HandleSendIfCharIsInChatMode(UserIndex)
        
    Case ClientPacketID.LoginExistingChar        'OLOGIN
        Call HandleLoginExistingChar(UserIndex)
        
    Case ClientPacketID.ThrowDices        'TIRDAD
        Call HandleThrowDices(UserIndex)
        
    Case ClientPacketID.LoginNewChar        'NLOGIN
        Call HandleLoginNewChar(UserIndex)
        
    Case ClientPacketID.DeleteChar
        Call HandleDeleteChar(UserIndex)
        
    Case ClientPacketID.Talk        ';
        Call HandleTalk(UserIndex)
        
    Case ClientPacketID.Yell        '-
        Call HandleYell(UserIndex)
        
    Case ClientPacketID.Whisper        '\
        Call HandleWhisper(UserIndex)
        
    Case ClientPacketID.Walk        'M
        Call HandleWalk(UserIndex)
        
    Case ClientPacketID.RequestPositionUpdate        'RPU
        Call HandleRequestPositionUpdate(UserIndex)
        
    Case ClientPacketID.Attack        'AT
        Call HandleAttack(UserIndex)
        
    Case ClientPacketID.PickUp        'AG
        Call HandlePickUp(UserIndex)
        
    Case ClientPacketID.SafeToggle        '/SEG & SEG  (SEG's behaviour has to be coded in the client)
        Call HandleSafeToggle(UserIndex)
        
    Case ClientPacketID.ResuscitationSafeToggle
        Call HandleResuscitationToggle(UserIndex)
        
    Case ClientPacketID.RequestGuildLeaderInfo        'GLINFO
        Call HandleRequestGuildLeaderInfo(UserIndex)
        
    Case ClientPacketID.RequestAtributes        'ATR
        Call HandleRequestAtributes(UserIndex)
        
    Case ClientPacketID.RequestFame        'FAMA
        Call HandleRequestFame(UserIndex)
        
    Case ClientPacketID.RequestSkills        'ESKI
        Call HandleRequestSkills(UserIndex)
        
    Case ClientPacketID.RequestMiniStats        'FEST
        Call HandleRequestMiniStats(UserIndex)
        
    Case ClientPacketID.CommerceEnd        'FINCOM
        Call HandleCommerceEnd(UserIndex)
        
    Case ClientPacketID.CommerceChat
        Call HandleCommerceChat(UserIndex)
        
    Case ClientPacketID.UserCommerceEnd        'FINCOMUSU
        Call HandleUserCommerceEnd(UserIndex)
        
    Case ClientPacketID.UserCommerceConfirm
        Call HandleUserCommerceConfirm(UserIndex)
        
    Case ClientPacketID.BankEnd        'FINBAN
        Call HandleBankEnd(UserIndex)
        
    Case ClientPacketID.UserCommerceOk        'COMUSUOK
        Call HandleUserCommerceOk(UserIndex)
        
    Case ClientPacketID.UserCommerceReject        'COMUSUNO
        Call HandleUserCommerceReject(UserIndex)
        
    Case ClientPacketID.Drop        'TI
        Call HandleDrop(UserIndex)
        
    Case ClientPacketID.CastSpell        'LH
        Call HandleCastSpell(UserIndex)
        
    Case ClientPacketID.LeftClick        'LC
        Call HandleLeftClick(UserIndex)
        
    Case ClientPacketID.DoubleClick        'RC
        Call HandleDoubleClick(UserIndex)
        
    Case ClientPacketID.Work        'UK
        Call HandleWork(UserIndex)
        
    Case ClientPacketID.UseSpellMacro        'UMH
        Call HandleUseSpellMacro(UserIndex)
        
    Case ClientPacketID.UseItem        'USA
        Call HandleUseItem(UserIndex)
        
    Case ClientPacketID.CraftBlacksmith        'CNS
        Call HandleCraftBlacksmith(UserIndex)
        
    Case ClientPacketID.CraftCarpenter        'CNC
        Call HandleCraftCarpenter(UserIndex)
        
    Case ClientPacketID.WorkLeftClick        'WLC
        Call HandleWorkLeftClick(UserIndex)
        
    Case ClientPacketID.CreateNewGuild        'CIG
        Call HandleCreateNewGuild(UserIndex)
        
    Case ClientPacketID.EquipItem        'EQUI
        Call HandleEquipItem(UserIndex)
        
    Case ClientPacketID.ChangeHeading        'CHEA
        Call HandleChangeHeading(UserIndex)
        
    Case ClientPacketID.ModifySkills        'SKSE
        Call HandleModifySkills(UserIndex)
        
    Case ClientPacketID.Train        'ENTR
        Call HandleTrain(UserIndex)
        
    Case ClientPacketID.CommerceBuy        'COMP
        Call HandleCommerceBuy(UserIndex)
        
    Case ClientPacketID.BankExtractItem        'RETI
        Call HandleBankExtractItem(UserIndex)
        
    Case ClientPacketID.CommerceSell        'VEND
        Call HandleCommerceSell(UserIndex)
        
    Case ClientPacketID.BankDeposit        'DEPO
        Call HandleBankDeposit(UserIndex)
        
    Case ClientPacketID.ForumPost        'DEMSG
        Call HandleForumPost(UserIndex)
        
    Case ClientPacketID.MoveSpell        'DESPHE
        Call HandleMoveSpell(UserIndex)
        
    Case ClientPacketID.MoveBank
        Call HandleMoveBank(UserIndex)
        
    Case ClientPacketID.ClanCodexUpdate        'DESCOD
        Call HandleClanCodexUpdate(UserIndex)
        
    Case ClientPacketID.UserCommerceOffer        'OFRECER
        Call HandleUserCommerceOffer(UserIndex)
        
    Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
        Call HandleGuildAcceptPeace(UserIndex)
        
    Case ClientPacketID.GuildRejectAlliance        'RECPALIA
        Call HandleGuildRejectAlliance(UserIndex)
        
    Case ClientPacketID.GuildRejectPeace        'RECPPEAT
        Call HandleGuildRejectPeace(UserIndex)
        
    Case ClientPacketID.GuildAcceptAlliance        'ACEPALIA
        Call HandleGuildAcceptAlliance(UserIndex)
        
    Case ClientPacketID.GuildOfferPeace        'PEACEOFF
        Call HandleGuildOfferPeace(UserIndex)
        
    Case ClientPacketID.GuildOfferAlliance        'ALLIEOFF
        Call HandleGuildOfferAlliance(UserIndex)
        
    Case ClientPacketID.GuildAllianceDetails        'ALLIEDET
        Call HandleGuildAllianceDetails(UserIndex)
        
    Case ClientPacketID.GuildPeaceDetails        'PEACEDET
        Call HandleGuildPeaceDetails(UserIndex)
        
    Case ClientPacketID.GuildRequestJoinerInfo        'ENVCOMEN
        Call HandleGuildRequestJoinerInfo(UserIndex)
        
    Case ClientPacketID.GuildAlliancePropList        'ENVALPRO
        Call HandleGuildAlliancePropList(UserIndex)
        
    Case ClientPacketID.GuildPeacePropList        'ENVPROPP
        Call HandleGuildPeacePropList(UserIndex)
        
    Case ClientPacketID.GuildDeclareWar        'DECGUERR
        Call HandleGuildDeclareWar(UserIndex)
        
    Case ClientPacketID.GuildNewWebsite        'NEWWEBSI
        Call HandleGuildNewWebsite(UserIndex)
        
    Case ClientPacketID.GuildAcceptNewMember        'ACEPTARI
        Call HandleGuildAcceptNewMember(UserIndex)
        
    Case ClientPacketID.GuildRejectNewMember        'RECHAZAR
        Call HandleGuildRejectNewMember(UserIndex)
        
    Case ClientPacketID.GuildKickMember        'ECHARCLA
        Call HandleGuildKickMember(UserIndex)
        
    Case ClientPacketID.GuildUpdateNews        'ACTGNEWS
        Call HandleGuildUpdateNews(UserIndex)
        
    Case ClientPacketID.GuildMemberInfo        '1HRINFO<
        Call HandleGuildMemberInfo(UserIndex)
        
    Case ClientPacketID.GuildOpenElections        'ABREELEC
        Call HandleGuildOpenElections(UserIndex)
        
    Case ClientPacketID.GuildRequestMembership        'SOLICITUD
        Call HandleGuildRequestMembership(UserIndex)
        
    Case ClientPacketID.GuildRequestDetails        'CLANDETAILS
        Call HandleGuildRequestDetails(UserIndex)
        
    Case ClientPacketID.Online        '/ONLINE
        Call HandleOnline(UserIndex)
        
    Case ClientPacketID.Quit        '/SALIR
        Call HandleQuit(UserIndex)
        
    Case ClientPacketID.GuildLeave        '/SALIRCLAN
        Call HandleGuildLeave(UserIndex)
        
    Case ClientPacketID.RequestAccountState        '/BALANCE
        Call HandleRequestAccountState(UserIndex)
        
    Case ClientPacketID.PetStand        '/QUIETO
        Call HandlePetStand(UserIndex)
        
    Case ClientPacketID.PetFollow        '/ACOMPANAR
        Call HandlePetFollow(UserIndex)
        
    Case ClientPacketID.ReleasePet        '/LIBERAR
        Call HandleReleasePet(UserIndex)
        
    Case ClientPacketID.TrainList        '/ENTRENAR
        Call HandleTrainList(UserIndex)
        
    Case ClientPacketID.Rest        '/DESCANSAR
        Call HandleRest(UserIndex)
        
    Case ClientPacketID.Meditate        '/MEDITAR
        Call HandleMeditate(UserIndex)
        
    Case ClientPacketID.Resucitate        '/RESUCITAR
        Call HandleResucitate(UserIndex)
        
    Case ClientPacketID.Heal        '/CURAR
        Call HandleHeal(UserIndex)
        
    Case ClientPacketID.Help        '/AYUDA
        Call HandleHelp(UserIndex)
        
    Case ClientPacketID.RequestStats        '/EST
        Call HandleRequestStats(UserIndex)
        
    Case ClientPacketID.CommerceStart        '/COMERCIAR
        Call HandleCommerceStart(UserIndex)
        
    Case ClientPacketID.BankStart        '/BOVEDA
        Call HandleBankStart(UserIndex)
        
    Case ClientPacketID.Enlist        '/ENLISTAR
        Call HandleEnlist(UserIndex)
        
    Case ClientPacketID.Information        '/INFORMACION
        Call HandleInformation(UserIndex)
        
    Case ClientPacketID.Reward        '/RECOMPENSA
        Call HandleReward(UserIndex)
        
    Case ClientPacketID.RequestMOTD        '/MOTD
        Call HandleRequestMOTD(UserIndex)
        
    Case ClientPacketID.UpTime        '/UPTIME
        Call HandleUpTime(UserIndex)
        
    Case ClientPacketID.PartyLeave        '/SALIRPARTY
        Call HandlePartyLeave(UserIndex)
        
    Case ClientPacketID.PartyCreate        '/CREARPARTY
        Call HandlePartyCreate(UserIndex)
        
    Case ClientPacketID.PartyJoin        '/PARTY
        Call HandlePartyJoin(UserIndex)
        
    Case ClientPacketID.Inquiry        '/ENCUESTA ( with no params )
        Call HandleInquiry(UserIndex)
        
    Case ClientPacketID.GuildMessage        '/CMSG
        Call HandleGuildMessage(UserIndex)
        
    Case ClientPacketID.PartyMessage        '/PMSG
        Call HandlePartyMessage(UserIndex)
        
    Case ClientPacketID.GuildOnline        '/ONLINECLAN
        Call HandleGuildOnline(UserIndex)
        
    Case ClientPacketID.PartyOnline        '/ONLINEPARTY
        Call HandlePartyOnline(UserIndex)
        
    Case ClientPacketID.CouncilMessage        '/BMSG
        Call HandleCouncilMessage(UserIndex)
        
    Case ClientPacketID.RoleMasterRequest        '/ROL
        Call HandleRoleMasterRequest(UserIndex)
        
    Case ClientPacketID.GMRequest        '/GM
        Call HandleGMRequest(UserIndex)
        
    Case ClientPacketID.bugReport        '/_BUG
        Call HandleBugReport(UserIndex)
        
    Case ClientPacketID.ChangeDescription        '/DESC
        Call HandleChangeDescription(UserIndex)
        
    Case ClientPacketID.GuildVote        '/VOTO
        Call HandleGuildVote(UserIndex)
        
    Case ClientPacketID.punishments        '/PENAS
        Call HandlePunishments(UserIndex)
        
    Case ClientPacketID.ChangePassword        '/CONTRASENA
        Call HandleChangePassword(UserIndex)
        
    Case ClientPacketID.Gamble        '/APOSTAR
        Call HandleGamble(UserIndex)
        
    Case ClientPacketID.InquiryVote        '/ENCUESTA ( with parameters )
        Call HandleInquiryVote(UserIndex)
        
    Case ClientPacketID.LeaveFaction        '/RETIRAR ( with no arguments )
        Call HandleLeaveFaction(UserIndex)
        
    Case ClientPacketID.BankExtractGold        '/RETIRAR ( with arguments )
        Call HandleBankExtractGold(UserIndex)
        
    Case ClientPacketID.BankDepositGold        '/DEPOSITAR
        Call HandleBankDepositGold(UserIndex)
        
    Case ClientPacketID.Denounce        '/DENUNCIAR
        Call HandleDenounce(UserIndex)
        
    Case ClientPacketID.GuildFundate        '/FUNDARCLAN
        Call HandleGuildFundate(UserIndex)
        
    Case ClientPacketID.GuildFundation
        Call HandleGuildFundation(UserIndex)
        
    Case ClientPacketID.PartyKick        '/ECHARPARTY
        Call HandlePartyKick(UserIndex)
        
    Case ClientPacketID.PartySetLeader        '/PARTYLIDER
        Call HandlePartySetLeader(UserIndex)
        
    Case ClientPacketID.PartyAcceptMember        '/ACCEPTPARTY
        Call HandlePartyAcceptMember(UserIndex)
        
    Case ClientPacketID.Ping        '/PING
        Call HandlePing(UserIndex)
        
    Case ClientPacketID.RequestPartyForm
        Call HandlePartyForm(UserIndex)
        
    Case ClientPacketID.ItemUpgrade
        Call HandleItemUpgrade(UserIndex)
        
    Case ClientPacketID.GMCommands        'GM Messages
        Call HandleGMCommands(UserIndex)
        
    Case ClientPacketID.InitCrafting
        Call HandleInitCrafting(UserIndex)
        
    Case ClientPacketID.Home
        Call HandleHome(UserIndex)
        
    Case ClientPacketID.ShowGuildNews
        Call HandleShowGuildNews(UserIndex)
        
    Case ClientPacketID.ShareNpc
        Call HandleShareNpc(UserIndex)
        
    Case ClientPacketID.StopSharingNpc
        Call HandleStopSharingNpc(UserIndex)
        
    Case ClientPacketID.Consultation
        Call HandleConsultation(UserIndex)
        
    Case ClientPacketID.moveItem
        Call HandleMoveItem(UserIndex)
        
    Case ClientPacketID.LoginExistingAccount
        Call HandleLoginExistingAccount(UserIndex)
        
    Case ClientPacketID.LoginNewAccount
        Call HandleLoginNewAccount(UserIndex)
        
    Case ClientPacketID.CentinelReport
        Call HandleCentinelReport(UserIndex)
        
    Case ClientPacketID.Ecvc
        Call HandleEnviaCvc(UserIndex)
        
    Case ClientPacketID.Acvc
        Call HandleAceptarCvc(UserIndex)
        
    Case ClientPacketID.IrCvc
        Call HandleIrCvc(UserIndex)
        
    Case ClientPacketID.DragAndDropHechizos
        Call HandleDragAndDropHechizos(UserIndex)
        
    Case ClientPacketID.Quest
        Call Quests.HandleQuest(UserIndex)
        
    Case ClientPacketID.QuestAccept
        Call Quests.HandleQuestAccept(UserIndex)
        
    Case ClientPacketID.QuestListRequest
        Call Quests.HandleQuestListRequest(UserIndex)
        
    Case ClientPacketID.QuestDetailsRequest
        Call Quests.HandleQuestDetailsRequest(UserIndex)
        
    Case ClientPacketID.QuestAbandon
        Call Quests.HandleQuestAbandon(UserIndex)
        
    Case ClientPacketID.CambiarContrasena
        Call HandleCambiarContrasena(UserIndex)
        
    Case ClientPacketID.FightSend
        Call HandleFightSend(UserIndex)
        
    Case ClientPacketID.FightAccept
        Call HandleFightAccept(UserIndex)
        
    Case ClientPacketID.CloseGuild
        Call HandleCloseGuild(UserIndex)
        
    Case ClientPacketID.Discord        '/Discord
        Call HandleDiscord(UserIndex)
        
    Case ClientPacketID.ObtenerDatosServer
        Call HandleObtenerDatosServer(UserIndex)
        
    Case ClientPacketID.CraftsmanCreate
        Call HandleCraftsmanCreate(UserIndex)
        
    Case ClientPacketID.AddAmigos
        Call Amigos.HandleAddAmigo(UserIndex)
        
    Case ClientPacketID.DelAmigos
        Call Amigos.HandleDelAmigo(UserIndex)
        
    Case ClientPacketID.OnAmigos
        Call Amigos.HandleOnAmigo(UserIndex)
        
    Case ClientPacketID.MsgAmigos
        Call Amigos.HandleMsgAmigo(UserIndex)
        
    Case ClientPacketID.LookProcess
        Call HandleLookProcess(UserIndex)
        
    Case ClientPacketID.SendProcessList
        Call HandleSendProcessList(UserIndex)
        
    Case Else
        'ERROR : Abort!
        Call CloseSocket(UserIndex)
        
End Select

If (Message.GetAvailable() > 0) Then
    Err.Raise &HDEADBEEF, "HandleIncomingData", "El paquete        '" & packetID & "' se encuentra en mal estado con '" & Message.GetAvailable() & "' bytes de mas por el usuario '" & UserList(UserIndex).Name & "'"
End If
Call PerformTimeLimitCheck(performance_timer, "Protocol handling message " & packetID)
HandleIncomingData_Err:

Set Reader = Nothing

If Err.Number <> 0 Then
    Debug.Print ("PackedID: " & packetID & vbNewLine & IIf(UserList(UserIndex).flags.UserLogged, "UserName: " & UserList(UserIndex).Name, "UserIndex: " & UserIndex) & " Protocol.HandleIncomingData")
    'Call CloseSocket(UserIndex)
    HandleIncomingData = False
End If

End Function

Public Sub WriteMultiMessage(ByVal UserIndex As Integer, _
       ByVal MessageIndex As Integer, _
       Optional ByVal Arg1 As Long, _
       Optional ByVal Arg2 As Long, _
       Optional ByVal Arg3 As Long, _
       Optional ByVal StringArg1 As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.MultiMessage)
    Call Writer.WriteInt8(MessageIndex)
    
    Select Case MessageIndex
        
        Case eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.SafeModeOn, eMessages.SafeModeOff, eMessages.ResuscitationSafeOff, eMessages.ResuscitationSafeOn, eMessages.NobilityLost, eMessages.CantUseWhileMeditating, eMessages.CancelHome, eMessages.FinishHome
            
        Case eMessages.NPCHitUser
            Call Writer.WriteInt8(Arg1)        'Target
            Call Writer.WriteInt16(Arg2)        'damage
            
        Case eMessages.UserHitNPC
            Call Writer.WriteInt32(Arg1)        'damage
            
        Case eMessages.UserAttackedSwing
            Call Writer.WriteInt16(UserList(Arg1).Char.CharIndex)
            
        Case eMessages.UserHittedByUser
            Call Writer.WriteInt16(Arg1)        'AttackerIndex
            Call Writer.WriteInt8(Arg2)        'Target
            Call Writer.WriteInt16(Arg3)        'damage
            
        Case eMessages.UserHittedUser
            Call Writer.WriteInt16(Arg1)        'AttackerIndex
            Call Writer.WriteInt8(Arg2)        'Target
            Call Writer.WriteInt16(Arg3)        'damage
            
        Case eMessages.WorkRequestTarget
            Call Writer.WriteInt8(Arg1)        'skill
            
        Case eMessages.HaveKilledUser        '"Has matado a " & UserList(VictimIndex).name & "!" "Has ganado " & DaExp & " puntos de experiencia."
            Call Writer.WriteInt16(UserList(Arg1).Char.CharIndex)        'VictimIndex
            Call Writer.WriteInt32(Arg2)        'Expe
            
        Case eMessages.UserKill        '"" & .name & " te ha matado!"
            Call Writer.WriteInt16(UserList(Arg1).Char.CharIndex)        'AttackerIndex
            
        Case eMessages.EarnExp
            Call Writer.WriteInt32(Arg1)
            
        Case eMessages.Home
            Call Writer.WriteInt8(CByte(Arg1))
            Call Writer.WriteInt16(CInt(Arg2))
            'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
            hasta que no se pasen los dats e .INFs al cliente, esto queda asi.
            Call Writer.WriteString8(StringArg1)        'Call Writer.WriteInt8(CByte(Arg2))
            
        Case eMessages.UserMuerto
            
        Case eMessages.NpcInmune
            
        Case eMessages.Hechizo_HechiceroMSG_NOMBRE
            Call Writer.WriteInt8(CByte(Arg1))        'SpellIndex
            Call Writer.WriteString8(StringArg1)        'Persona
            
        Case eMessages.Hechizo_HechiceroMSG_ALGUIEN
            Call Writer.WriteInt8(CByte(Arg1))        'SpellIndex
            
        Case eMessages.Hechizo_HechiceroMSG_CRIATURA
            Call Writer.WriteInt8(CByte(Arg1))        'SpellIndex
            
        Case eMessages.Hechizo_PropioMSG
            Call Writer.WriteInt8(CByte(Arg1))        'SpellIndex
            
        Case eMessages.Hechizo_TargetMSG
            Call Writer.WriteInt8(CByte(Arg1))        'SpellIndex
            Call Writer.WriteString8(StringArg1)        'Persona
            
    End Select
    
    Exit Sub        ''
    
errHandler:
    
    Call Writer.Clear
    
End Sub

Private Sub HandleGMCommands(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    Dim Command     As Long
    
    With UserList(UserIndex)
        
        Command = CLng(Reader.ReadInt8)
        
        Select Case Command
            
            Case eGMCommands.GMMessage        '/GMSG
                Call HandleGMMessage(UserIndex)
                
            Case eGMCommands.showName        '/SHOWNAME
                Call HandleShowName(UserIndex)
                
            Case eGMCommands.OnlineRoyalArmy
                Call HandleOnlineRoyalArmy(UserIndex)
                
            Case eGMCommands.OnlineChaosLegion        '/ONLINECAOS
                Call HandleOnlineChaosLegion(UserIndex)
                
            Case eGMCommands.GoNearby        '/IRCERCA
                Call HandleGoNearby(UserIndex)
                
            Case eGMCommands.comment        '/REM
                Call HandleComment(UserIndex)
                
            Case eGMCommands.serverTime        '/HORA
                Call HandleServerTime(UserIndex)
                
            Case eGMCommands.Where        '/DONDE
                Call HandleWhere(UserIndex)
                
            Case eGMCommands.CreaturesInMap        '/NENE
                Call HandleCreaturesInMap(UserIndex)
                
            Case eGMCommands.WarpMeToTarget        '/TELEPLOC
                Call HandleWarpMeToTarget(UserIndex)
                
            Case eGMCommands.WarpChar        '/TELEP
                Call HandleWarpChar(UserIndex)
                
            Case eGMCommands.Silence        '/SILENCIAR
                Call HandleSilence(UserIndex)
                
            Case eGMCommands.SOSShowList        '/SHOW SOS
                Call HandleSOSShowList(UserIndex)
                
            Case eGMCommands.SOSRemove        'SOSDONE
                Call HandleSOSRemove(UserIndex)
                
            Case eGMCommands.GoToChar        '/IRA
                Call HandleGoToChar(UserIndex)
                
            Case eGMCommands.invisible        '/INVISIBLE
                Call HandleInvisible(UserIndex)
                
            Case eGMCommands.GMPanel        '/PANELGM
                Call HandleGMPanel(UserIndex)
                
            Case eGMCommands.RequestUserList        'LISTUSU
                Call HandleRequestUserList(UserIndex)
                
            Case eGMCommands.Working        '/TRABAJANDO
                Call HandleWorking(UserIndex)
                
            Case eGMCommands.Hiding        '/OCULTANDO
                Call HandleHiding(UserIndex)
                
            Case eGMCommands.Jail        '/CARCEL
                Call HandleJail(UserIndex)
                
            Case eGMCommands.KillNPC        '/RMATA
                Call HandleKillNPC(UserIndex)
                
            Case eGMCommands.WarnUser        '/ADVERTENCIA
                Call HandleWarnUser(UserIndex)
                
            Case eGMCommands.EditChar        '/MOD
                Call HandleEditChar(UserIndex)
                
            Case eGMCommands.RequestCharInfo        '/INFO
                Call HandleRequestCharInfo(UserIndex)
                
            Case eGMCommands.RequestCharStats        '/STAT
                Call HandleRequestCharStats(UserIndex)
                
            Case eGMCommands.RequestCharGold        '/BAL
                Call HandleRequestCharGold(UserIndex)
                
            Case eGMCommands.RequestCharInventory        '/INV
                Call HandleRequestCharInventory(UserIndex)
                
            Case eGMCommands.RequestCharBank        '/BOV
                Call HandleRequestCharBank(UserIndex)
                
            Case eGMCommands.RequestCharSkills        '/SKILLS
                Call HandleRequestCharSkills(UserIndex)
                
            Case eGMCommands.ReviveChar        '/REVIVIR
                Call HandleReviveChar(UserIndex)
                
            Case eGMCommands.OnlineGM        '/ONLINEGM
                Call HandleOnlineGM(UserIndex)
                
            Case eGMCommands.OnlineMap        '/ONLINEMAP
                Call HandleOnlineMap(UserIndex)
                
            Case eGMCommands.Forgive        '/PERDON
                Call HandleForgive(UserIndex)
                
            Case eGMCommands.Kick        '/ECHAR
                Call HandleKick(UserIndex)
                
            Case eGMCommands.Execute        '/EJECUTAR
                Call HandleExecute(UserIndex)
                
            Case eGMCommands.BanChar        '/BAN
                Call HandleBanChar(UserIndex)
                
            Case eGMCommands.UnbanChar        '/UNBAN
                Call HandleUnbanChar(UserIndex)
                
            Case eGMCommands.NPCFollow        '/SEGUIR
                Call HandleNPCFollow(UserIndex)
                
            Case eGMCommands.SummonChar        '/SUM
                Call HandleSummonChar(UserIndex)
                
            Case eGMCommands.SpawnListRequest        '/CC
                Call HandleSpawnListRequest(UserIndex)
                
            Case eGMCommands.SpawnCreature        'SPA
                Call HandleSpawnCreature(UserIndex)
                
            Case eGMCommands.ResetNPCInventory        '/RESETINV
                Call HandleResetNPCInventory(UserIndex)
                
            Case eGMCommands.ServerMessage        '/RMSG
                Call HandleServerMessage(UserIndex)
                
            Case eGMCommands.MapMessage        '/MAPMSG
                Call HandleMapMessage(UserIndex)
                
            Case eGMCommands.NickToIP        '/NICK2IP
                Call HandleNickToIP(UserIndex)
                
            Case eGMCommands.IPToNick        '/IP2NICK
                Call HandleIPToNick(UserIndex)
                
            Case eGMCommands.GuildOnlineMembers        '/ONCLAN
                Call HandleGuildOnlineMembers(UserIndex)
                
            Case eGMCommands.TeleportCreate        '/CT
                Call HandleTeleportCreate(UserIndex)
                
            Case eGMCommands.TeleportDestroy        '/DT
                Call HandleTeleportDestroy(UserIndex)
                
            Case eGMCommands.RainToggle        '/LLUVIA
                Call HandleRainToggle(UserIndex)
                
            Case eGMCommands.SetCharDescription        '/SETDESC
                Call HandleSetCharDescription(UserIndex)
                
            Case eGMCommands.ForceMP3ToMap        '/FORCEMP3MAP
                Call HanldeForceMP3ToMap(UserIndex)
                
            Case eGMCommands.ForceMIDIToMap        '/FORCEMIDIMAP
                Call HanldeForceMIDIToMap(UserIndex)
                
            Case eGMCommands.ForceWAVEToMap        '/FORCEWAVMAP
                Call HandleForceWAVEToMap(UserIndex)
                
            Case eGMCommands.RoyalArmyMessage        '/REALMSG
                Call HandleRoyalArmyMessage(UserIndex)
                
            Case eGMCommands.ChaosLegionMessage        '/CAOSMSG
                Call HandleChaosLegionMessage(UserIndex)
                
            Case eGMCommands.CitizenMessage        '/CIUMSG
                Call HandleCitizenMessage(UserIndex)
                
            Case eGMCommands.CriminalMessage        '/CRIMSG
                Call HandleCriminalMessage(UserIndex)
                
            Case eGMCommands.TalkAsNPC        '/TALKAS
                Call HandleTalkAsNPC(UserIndex)
                
            Case eGMCommands.DestroyAllItemsInArea        '/MASSDEST
                Call HandleDestroyAllItemsInArea(UserIndex)
                
            Case eGMCommands.AcceptRoyalCouncilMember        '/ACEPTCONSE
                Call HandleAcceptRoyalCouncilMember(UserIndex)
                
            Case eGMCommands.AcceptChaosCouncilMember        '/ACEPTCONSECAOS
                Call HandleAcceptChaosCouncilMember(UserIndex)
                
            Case eGMCommands.ItemsInTheFloor        '/PISO
                Call HandleItemsInTheFloor(UserIndex)
                
            Case eGMCommands.MakeDumb        '/ESTUPIDO
                Call HandleMakeDumb(UserIndex)
                
            Case eGMCommands.MakeDumbNoMore        '/NOESTUPIDO
                Call HandleMakeDumbNoMore(UserIndex)
                
            Case eGMCommands.DumpIPTables        '/DUMPSECURITY
                Call HandleDumpIPTables(UserIndex)
                
            Case eGMCommands.CouncilKick        '/KICKCONSE
                Call HandleCouncilKick(UserIndex)
                
            Case eGMCommands.SetTrigger        '/TRIGGER
                Call HandleSetTrigger(UserIndex)
                
            Case eGMCommands.AskTrigger        '/TRIGGER with no args
                Call HandleAskTrigger(UserIndex)
                
            Case eGMCommands.BannedIPList        '/BANIPLIST
                Call HandleBannedIPList(UserIndex)
                
            Case eGMCommands.BannedIPReload        '/BANIPRELOAD
                Call HandleBannedIPReload(UserIndex)
                
            Case eGMCommands.GuildMemberList        '/MIEMBROSCLAN
                Call HandleGuildMemberList(UserIndex)
                
            Case eGMCommands.GuildBan        '/BANCLAN
                Call HandleGuildBan(UserIndex)
                
            Case eGMCommands.BanIP        '/BANIP
                Call HandleBanIP(UserIndex)
                
            Case eGMCommands.UnbanIP        '/UNBANIP
                Call HandleUnbanIP(UserIndex)
                
            Case eGMCommands.CreateItem        '/CI
                Call HandleCreateItem(UserIndex)
                
            Case eGMCommands.DestroyItems        '/DEST
                Call HandleDestroyItems(UserIndex)
                
            Case eGMCommands.ChaosLegionKick        '/NOCAOS
                Call HandleChaosLegionKick(UserIndex)
                
            Case eGMCommands.RoyalArmyKick        '/NOREAL
                Call HandleRoyalArmyKick(UserIndex)
                
            Case eGMCommands.ForceMP3All        '/FORCEMP3
                Call HandleForceMP3All(UserIndex)
                
            Case eGMCommands.ForceMIDIAll        '/FORCEMIDI
                Call HandleForceMIDIAll(UserIndex)
                
            Case eGMCommands.ForceWAVEAll        '/FORCEWAV
                Call HandleForceWAVEAll(UserIndex)
                
            Case eGMCommands.RemovePunishment        '/BORRARPENA
                Call HandleRemovePunishment(UserIndex)
                
            Case eGMCommands.TileBlockedToggle        '/BLOQ
                Call HandleTileBlockedToggle(UserIndex)
                
            Case eGMCommands.KillNPCNoRespawn        '/MATA
                Call HandleKillNPCNoRespawn(UserIndex)
                
            Case eGMCommands.KillAllNearbyNPCs        '/MASSKILL
                Call HandleKillAllNearbyNPCs(UserIndex)
                
            Case eGMCommands.LastIP        '/LASTIP
                Call HandleLastIP(UserIndex)
                
            Case eGMCommands.ChangeMOTD        '/MOTDCAMBIA
                Call HandleChangeMOTD(UserIndex)
                
            Case eGMCommands.SetMOTD        'ZMOTD
                Call HandleSetMOTD(UserIndex)
                
            Case eGMCommands.SystemMessage        '/SMSG
                Call HandleSystemMessage(UserIndex)
                
            Case eGMCommands.CreateNPC        '/ACC y /RACC
                Call HandleCreateNPC(UserIndex)
                
            Case eGMCommands.ImperialArmour        '/AI1 - 4
                Call HandleImperialArmour(UserIndex)
                
            Case eGMCommands.ChaosArmour        '/AC1 - 4
                Call HandleChaosArmour(UserIndex)
                
            Case eGMCommands.NavigateToggle        '/NAVE
                Call HandleNavigateToggle(UserIndex)
                
            Case eGMCommands.ServerOpenToUsersToggle        '/HABILITAR
                Call HandleServerOpenToUsersToggle(UserIndex)
                
            Case eGMCommands.TurnOffServer        '/APAGAR
                Call HandleTurnOffServer(UserIndex)
                
            Case eGMCommands.TurnCriminal        '/CONDEN
                Call HandleTurnCriminal(UserIndex)
                
            Case eGMCommands.ResetFactions        '/RAJAR
                Call HandleResetFactions(UserIndex)
                
            Case eGMCommands.RemoveCharFromGuild        '/RAJARCLAN
                Call HandleRemoveCharFromGuild(UserIndex)
                
            Case eGMCommands.RequestCharMail        '/LASTEMAIL
                Call HandleRequestCharMail(UserIndex)
                
            Case eGMCommands.AlterPassword        '/APASS
                Call HandleAlterPassword(UserIndex)
                
            Case eGMCommands.AlterMail        '/AEMAIL
                Call HandleAlterMail(UserIndex)
                
            Case eGMCommands.AlterName        '/ANAME
                Call HandleAlterName(UserIndex)
                
            Case Declaraciones.eGMCommands.DoBackUp        '/DOBACKUP
                Call HandleDoBackUp(UserIndex)
                
            Case eGMCommands.ShowGuildMessages        '/SHOWCMSG
                Call HandleShowGuildMessages(UserIndex)
                
            Case eGMCommands.SaveMap        '/GUARDAMAPA
                Call HandleSaveMap(UserIndex)
                
            Case eGMCommands.ChangeMapInfoPK        '/MODMAPINFO PK
                Call HandleChangeMapInfoPK(UserIndex)
                
            Case eGMCommands.ChangeMapInfoBackup        '/MODMAPINFO BACKUP
                Call HandleChangeMapInfoBackup(UserIndex)
                
            Case eGMCommands.ChangeMapInfoRestricted        '/MODMAPINFO RESTRINGIR
                Call HandleChangeMapInfoRestricted(UserIndex)
                
            Case eGMCommands.ChangeMapInfoNoMagic        '/MODMAPINFO MAGIASINEFECTO
                Call HandleChangeMapInfoNoMagic(UserIndex)
                
            Case eGMCommands.ChangeMapInfoNoInvi        '/MODMAPINFO INVISINEFECTO
                Call HandleChangeMapInfoNoInvi(UserIndex)
                
            Case eGMCommands.ChangeMapInfoNoResu        '/MODMAPINFO RESUSINEFECTO
                Call HandleChangeMapInfoNoResu(UserIndex)
                
            Case eGMCommands.ChangeMapInfoLand        '/MODMAPINFO TERRENO
                Call HandleChangeMapInfoLand(UserIndex)
                
            Case eGMCommands.ChangeMapInfoZone        '/MODMAPINFO ZONA
                Call HandleChangeMapInfoZone(UserIndex)
                
            Case eGMCommands.ChangeMapInfoStealNpc        '/MODMAPINFO ROBONPC
                Call HandleChangeMapInfoStealNpc(UserIndex)
                
            Case eGMCommands.ChangeMapInfoNoOcultar        '/MODMAPINFO OCULTARSINEFECTO
                Call HandleChangeMapInfoNoOcultar(UserIndex)
                
            Case eGMCommands.ChangeMapInfoNoInvocar        '/MODMAPINFO INVOCARSINEFECTO
                Call HandleChangeMapInfoNoInvocar(UserIndex)
                
            Case eGMCommands.SaveChars        '/GRABAR
                Call HandleSaveChars(UserIndex)
                
            Case eGMCommands.CleanSOS        '/BORRAR SOS
                Call HandleCleanSOS(UserIndex)
                
            Case eGMCommands.ShowServerForm        '/SHOW INT
                Call HandleShowServerForm(UserIndex)
                
            Case eGMCommands.night        '/NOCHE
                Call HandleNight(UserIndex)
                
            Case eGMCommands.KickAllChars        '/ECHARTODOSPJS
                Call HandleKickAllChars(UserIndex)
                
            Case eGMCommands.ReloadNPCs        '/RELOADNPCS
                Call HandleReloadNPCs(UserIndex)
                
            Case eGMCommands.ReloadServerIni        '/RELOADSINI
                Call HandleReloadServerIni(UserIndex)
                
            Case eGMCommands.ReloadSpells        '/RELOADHECHIZOS
                Call HandleReloadSpells(UserIndex)
                
            Case eGMCommands.ReloadObjects        '/RELOADOBJ
                Call HandleReloadObjects(UserIndex)
                
            Case eGMCommands.Restart        '/REINICIAR
                Call HandleRestart(UserIndex)
                
            Case eGMCommands.ResetAutoUpdate        '/AUTOUPDATE
                Call HandleResetAutoUpdate(UserIndex)
                
            Case eGMCommands.ChatColor        '/CHATCOLOR
                Call HandleChatColor(UserIndex)
                
            Case eGMCommands.Ignored        '/IGNORADO
                Call HandleIgnored(UserIndex)
                
            Case eGMCommands.CheckSlot        '/SLOT
                Call HandleCheckSlot(UserIndex)
                
            Case eGMCommands.SetIniVar        '/SETINIVAR LLAVE CLAVE VALOR
                Call HandleSetIniVar(UserIndex)
                
            Case eGMCommands.CreatePretorianClan        '/CREARPRETORIANOS
                Call HandleCreatePretorianClan(UserIndex)
                
            Case eGMCommands.RemovePretorianClan        '/ELIMINARPRETORIANOS
                Call HandleDeletePretorianClan(UserIndex)
                
            Case eGMCommands.EnableDenounces        '/DENUNCIAS
                Call HandleEnableDenounces(UserIndex)
                
            Case eGMCommands.ShowDenouncesList        '/SHOW DENUNCIAS
                Call HandleShowDenouncesList(UserIndex)
                
            Case eGMCommands.SetDialog        '/SETDIALOG
                Call HandleSetDialog(UserIndex)
                
            Case eGMCommands.Impersonate        '/IMPERSONAR
                Call HandleImpersonate(UserIndex)
                
            Case eGMCommands.Imitate        '/MIMETIZAR
                Call HandleImitate(UserIndex)
                
            Case eGMCommands.RecordAdd
                Call HandleRecordAdd(UserIndex)
                
            Case eGMCommands.RecordAddObs
                Call HandleRecordAddObs(UserIndex)
                
            Case eGMCommands.RecordRemove
                Call HandleRecordRemove(UserIndex)
                
            Case eGMCommands.RecordListRequest
                Call HandleRecordListRequest(UserIndex)
                
            Case eGMCommands.RecordDetailsRequest
                Call HandleRecordDetailsRequest(UserIndex)
                
            Case eGMCommands.ExitDestroy
                Call HandleExitDestroy(UserIndex)
                
            Case eGMCommands.ToggleCentinelActivated        '/CENTINELAACTIVADO
                Call HandleToggleCentinelActivated(UserIndex)
                
            Case eGMCommands.SearchNpc        '/BUSCAR
                Call HandleSearchNpc(UserIndex)
                
            Case eGMCommands.SearchObj        '/BUSCAR
                Call HandleSearchObj(UserIndex)
                
            Case eGMCommands.LimpiarMundo        '/LIMPIARMUNDO
                Call HandleLimpiarMundo(UserIndex)
                
        End Select
        
    End With
    
    Exit Sub
    
errHandler:
    Call LogError("Error en GmCommands. Error: " & Err.Number & " - " & Err.description & ". Paquete: " & Command)
    
End Sub

''
' Handles the "Home" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleHome(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Budi
    'Creation Date: 06/01/2010
    'Last Modification: 05/06/10
    'Pato - 05/06/10: Add the Ucase$ to prevent problems.
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.TargetNpcTipo = eNPCType.Gobernador Then
            Call setHome(UserIndex, Npclist(.flags.TargetNPC).Ciudad, .flags.TargetNPC)
        Else
            
            If .flags.Muerto = 1 Then
                
                'Si es un mapa comUn y no esta en cana
                If (MapInfo(.Pos.Map).Restringir = eRestrict.restrict_no) And (.Counters.Pena = 0) Then
                    If .flags.Traveling = 0 Then
                        If Ciudades(.Hogar).Map <> .Pos.Map Then
                            Call goHome(UserIndex)
                        Else
                            Call WriteConsoleMsg(UserIndex, "Ya te encuentras en tu hogar.", FontTypeNames.FONTTYPE_INFO)
                            
                        End If
                        
                    Else
                        Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
                        .flags.Traveling = 0
                        .Counters.goHome = 0
                        
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes usar este comando aqui.", FontTypeNames.FONTTYPE_FIGHT)
                    
                End If
                
            Else
                Call WriteConsoleMsg(UserIndex, "Debes estar muerto para utilizar este comando.", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        End If
        
    End With
    
End Sub

''
' Handles the "DeleteChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDeleteChar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Recoaro (Recox)
    'Last Modification: 07/01/20
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    Dim UserName    As String
    Dim AccountHash As String
    UserName = Reader.ReadString8()
    AccountHash = Reader.ReadString8()
    
    'If we got here then packet is complete, copy data back to original queue
    
    Call BorrarUsuario(UserIndex, UserName, AccountHash)
    
    'Enviamos paquete para mostrar mensaje satisfactorio en el cliente
    Call Writer.WriteInt8(ServerPacketID.DeletedChar)
    
    Exit Sub
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Writer.Clear
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "LoginExistingChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo errHandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    
    Dim UserName    As String
    Dim AccountHash As String
    Dim version     As String
    
    UserName = Reader.ReadString8()
    AccountHash = Reader.ReadString8()
    
    'Convert version number to string
    version = CStr(Reader.ReadInt8) & "." & CStr(Reader.ReadInt8) & "." & CStr(Reader.ReadInt8)
    
    'If we got here then packet is complete
    
    If Not AsciiValidos(UserName) Then
        Call WriteErrorMsg(UserIndex, "Nombre invalido.")
        Call CloseSocket(UserIndex)
        
        Exit Sub
        
    End If
    
    If Not PersonajeExiste(UserName) Then
        Call WriteErrorMsg(UserIndex, "El personaje no existe.")
        Call CloseSocket(UserIndex)
        
        Exit Sub
        
    End If
    
    If BANCheck(UserName) Then
        Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a Argentum Online debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.argentumonline.org")
    ElseIf Not VersionOK(version) Then
        Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la ultima version es la " & ULTIMAVERSION & ". Tu Version " & version & ". La misma se encuentra disponible en www.argentumonline.org")
    Else
        Call ConnectUser(UserIndex, UserName, AccountHash)
    End If
    
errHandler:
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
End Sub

''
' Handles the "ThrowDices" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleThrowDices(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex).Stats
        .UserAtributos(eAtributos.Fuerza) = RandomNumber(DiceMinimum, DiceMaximum)
        .UserAtributos(eAtributos.Agilidad) = RandomNumber(DiceMinimum, DiceMaximum)
        .UserAtributos(eAtributos.Inteligencia) = RandomNumber(DiceMinimum, DiceMaximum)
        .UserAtributos(eAtributos.Carisma) = RandomNumber(DiceMinimum, DiceMaximum)
        .UserAtributos(eAtributos.Constitucion) = RandomNumber(DiceMinimum, DiceMaximum)
        
    End With
    
    Call WriteDiceRoll(UserIndex)
    
End Sub

''
' Handles the "LoginNewChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    'This packet contains strings
    
    Dim UserName    As String
    Dim AccountHash As String
    Dim version     As String
    Dim race        As eRaza
    Dim gender      As eGenero
    Dim homeland    As eCiudad
    Dim Class       As eClass
    Dim Head        As Integer
    
    UserName = Reader.ReadString8()
    AccountHash = Reader.ReadString8()
    
    'Convert version number to string
    version = CStr(Reader.ReadInt8) & "." & CStr(Reader.ReadInt8) & "." & CStr(Reader.ReadInt8)
    
    race = Reader.ReadInt8
    gender = Reader.ReadInt8
    Class = Reader.ReadInt8
    Head = Reader.ReadInt16
    homeland = Reader.ReadInt8
    
    'If we got here then packet is complete
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(UserIndex, "La creacion de personajes en este servidor se ha deshabilitado.")
        Call CloseSocket(UserIndex)
        
        Exit Sub
        
    End If
    
    If ServerSoloGMs <> 0 Then
        Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Consulte la pagina oficial o el foro oficial para mas informacion.")
        Call CloseSocket(UserIndex)
        
        Exit Sub
        
    End If
    
    If aClon.MaxPersonajes(UserList(UserIndex).IP) Then
        Call WriteErrorMsg(UserIndex, "Has creado demasiados personajes.")
        Call CloseSocket(UserIndex)
        
        Exit Sub
        
    End If
    
    If Not VersionOK(version) Then
        Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la ultima version es la " & ULTIMAVERSION & ". Tu Version " & version & ". La misma se encuentra disponible en www.argentumonline.org")
    Else
        Call ConnectNewUser(UserIndex, UserName, AccountHash, race, gender, Class, homeland, Head)
        
    End If
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010
    '15/07/2009: ZaMa - Now invisible admins talk by console.
    '23/09/2009: ZaMa - Now invisible admins can't send empty chat.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Chat    As String
        
        Chat = Reader.ReadString8()
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Dijo: " & Chat)
            
        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .Clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
                    
                End If
                
            Else
                
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If Not (.flags.AdminInvisible = 1) Then
                If .flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, .flags.ChatColor))
                    
                End If
                
            Else
                
                If RTrim(Chat) <> "" Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010 (ZaMa)
    '15/07/2009: ZaMa - Now invisible admins yell by console.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Chat    As String
        
        Chat = Reader.ReadString8()
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Grito: " & Chat)
            
        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .Clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
                    
                End If
                
            Else
                
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .flags.Privilegios And PlayerType.User Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, vbRed))
                    
                End If
                
            Else
                
                If Not (.flags.AdminInvisible = 1) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_GM_YELL))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 03/12/2010
    '28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
    '15/07/2009: ZaMa - Now invisible admins wisper by console.
    '03/12/2010: Enanoh - Agregue susurro a Admins en modo consulta y Los Dioses pueden susurrar en ciertos casos.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Chat            As String
        
        Dim TargetUserIndex As Integer
        
        Dim TargetPriv      As PlayerType
        
        Dim UserPriv        As PlayerType
        
        Dim TargetName      As String
        
        TargetName = Reader.ReadString8()
        Chat = Reader.ReadString8()
        
        UserPriv = .flags.Privilegios
        
        If .flags.Muerto Then
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
        Else
            ' Offline
            TargetUserIndex = NameIndex(TargetName)
            
            If TargetUserIndex = INVALID_INDEX Then
                
                ' Admin
                If EsGmChar(TargetName) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                    ' Whisperer admin? (Else say nothing)
                ElseIf (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
                ' Online
            Else
                ' Privilegios
                TargetPriv = UserList(TargetUserIndex).flags.Privilegios
                
                ' Consejeros, semis y usuarios no pueden susurrar a dioses (Salvo en consulta)
                If (TargetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (UserPriv And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios)) <> 0 And Not .flags.EnConsulta Then
                    
                    ' No puede
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                    
                    ' Usuarios no pueden susurrar a semis o conses (Salvo en consulta)
                ElseIf (UserPriv And PlayerType.User) <> 0 And (Not TargetPriv And PlayerType.User) <> 0 And Not .flags.EnConsulta Then
                    
                    ' No puede
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                    
                    ' En rango? (Los dioses pueden susurrar a distancia)
                ElseIf Not EstaPCarea(UserIndex, TargetUserIndex) And (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) = 0 Then
                    
                    ' No se puede susurrar a admins fuera de su rango
                    If (TargetPriv And (PlayerType.User)) = 0 And (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) = 0 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                        
                        ' Whisperer admin? (Else say nothing)
                    ElseIf (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                        Call WriteConsoleMsg(UserIndex, "Estas muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                Else
                    
                    '[Consejeros & GMs]
                    If UserPriv And (PlayerType.consejero Or PlayerType.SemiDios) Then
                        Call LogGM(.Name, "Le susurro a        '" & UserList(TargetUserIndex).Name & "' " & Chat)
                        
                        ' Usuarios a administradores
                    ElseIf (UserPriv And PlayerType.User) <> 0 And (TargetPriv And PlayerType.User) = 0 Then
                        Call LogGM(UserList(TargetUserIndex).Name, .Name & " le susurro en consulta: " & Chat)
                        
                    End If
                    
                    If LenB(Chat) <> 0 Then
                        'Analize chat...
                        Call Statistics.ParseChat(Chat)
                        
                        ' Dios susurrando a distancia
                        If Not EstaPCarea(UserIndex, TargetUserIndex) And (UserPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                            
                            Call WriteConsoleMsg(UserIndex, "Susurraste> " & Chat, FontTypeNames.FONTTYPE_GM)
                            Call WriteConsoleMsg(TargetUserIndex, "Gm susurra> " & Chat, FontTypeNames.FONTTYPE_GM)
                            
                        ElseIf Not (.flags.AdminInvisible = 1) Then
                            Call WriteChatOverHead(UserIndex, Chat, .Char.CharIndex, vbBlue)
                            Call WriteChatOverHead(TargetUserIndex, Chat, .Char.CharIndex, vbBlue)
                            
                            '[CDT 17-02-2004]
                            If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageChatOverHead("A " & UserList(TargetUserIndex).Name & "> " & Chat, .Char.CharIndex, vbYellow))
                                
                            End If
                            
                        Else
                            Call WriteConsoleMsg(UserIndex, "Susurraste> " & Chat, FontTypeNames.FONTTYPE_GM)
                            
                            If UserIndex <> TargetUserIndex Then Call WriteConsoleMsg(TargetUserIndex, "Gm susurra> " & Chat, FontTypeNames.FONTTYPE_GM)
                            
                            If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageConsoleMsg("Gm dijo a " & UserList(TargetUserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_GM))
                                
                            End If
                            
                        End If
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/01/2012 (Recox)
    '11/19/09 Pato - Now the class bandit can walk hidden.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '12/01/2020: Recox - TiempoDeWalk agregado para las monturas
    '***************************************************
    
    Dim dummy       As Long
    
    Dim TempTick    As Long
    
    Dim heading     As eHeading
    
    With UserList(UserIndex)
        
        heading = Reader.ReadInt8
        
        Dim TiempoDeWalk As Byte
        If .flags.Equitando = 1 Then
            TiempoDeWalk = 36
        Else
            TiempoDeWalk = 30
        End If
        
        'Prevent SpeedHack
        If .flags.TimesWalk >= TiempoDeWalk Then
            TempTick = GetTickCount And &H7FFFFFFF
            dummy = (TempTick - .flags.StartWalk)
            
            ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
            '(it's about 193 ms per step against the over 200 needed in perfect conditions)
            If dummy < 5800 Then
                If TempTick - .flags.CountSH > 30000 Then
                    .flags.CountSH = 0
                    
                End If
                
                If Not .flags.CountSH = 0 Then
                    If dummy <> 0 Then dummy = 126000 \ dummy
                    
                    Call LogHackAttemp("Tramposo SH: " & .Name & " , " & dummy)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha sido echado por el servidor por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER))
                    Call CloseSocket(UserIndex)
                    
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
        Call CancelExit(UserIndex)
        
        If .flags.Comerciando Then Exit Sub
        
        'Esta usando el /HOGAR, no se puede mover
        If .flags.Traveling = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes moverte mientras estas viajando a tu hogar con el comando /HOGAR.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
                .Char.loops = 0
                
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
                Call MoveUserChar(UserIndex, heading)
            Else
                'Move user
                Call MoveUserChar(UserIndex, heading)
                
                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                    
                    Call WriteRestOK(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        Else        'paralized
            
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteConsoleMsg(UserIndex, "No puedes moverte porque estas paralizado.", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
            .flags.CountSH = 0
            
        End If
        
        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
            If .Clase <> eClass.Thief And .Clase <> eClass.Bandit Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                If .flags.Navegando = 1 Then
                    If .Clase = eClass.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
                        Call ToggleBoatBody(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
                        
                    End If
                    
                Else
                    
                    'If not under a spell effect, show char
                    If .flags.invisible = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                        Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call WritePosUpdate(UserIndex)
    
End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010
    'Last Modified By: ZaMa
    '10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo.
    '13/11/2009: ZaMa - Se cancela el estado no atacable al atcar.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    With UserList(UserIndex)
        
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'If user meditates, can't attack
        If .flags.Meditando Then
            Exit Sub
            
        End If
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes usar asi este arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Play AttackAnim on Clients
        Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterAttackAnim(.Char.CharIndex))
        
        'Attack!
        Call UsuarioAtaca(UserIndex)
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .Clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
                    
                End If
                
            Else
                
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/25/09
    '02/26/2006: Marco - Agregue un checkeo por si el usuario trata de agarrar un item mientras comercia.
    '***************************************************
    With UserList(UserIndex)
        
        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then Exit Sub
        
        'If user is trading items and attempts to pickup an item, he's cheating, so we kick him.
        If .flags.Comerciando Then Exit Sub
        
        'Lower rank administrators can't pick up items
        If .flags.Privilegios And PlayerType.consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteConsoleMsg(UserIndex, "No puedes tomar ningUn objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
        End If
        
        Call GetObj(UserIndex)
        
    End With
    
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Seguro Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff)        'Call WriteSafeModeOff(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn)        'Call WriteSafeModeOn(UserIndex)
            
        End If
        
        .flags.Seguro = Not .flags.Seguro
        
    End With
    
End Sub

''
' Handles the "ResuscitationSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResuscitationToggle(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Rapsodius
    'Creation Date: 10/10/07
    '***************************************************
    With UserList(UserIndex)
        
        .flags.SeguroResu = Not .flags.SeguroResu
        
        If .flags.SeguroResu Then
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn)        'Call WriteResuscitationSafeOn(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff)        'Call WriteResuscitationSafeOff(UserIndex)
            
        End If
        
    End With
    
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call modGuilds.SendGuildLeaderInfo(UserIndex)
    
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call WriteAttributes(UserIndex)
    
End Sub

''
' Handles the "RequestFame" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestFame(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call EnviarFama(UserIndex)
    
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call WriteSendSkills(UserIndex)
    
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call WriteMiniStats(UserIndex)
    
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    'User quits commerce mode
    UserList(UserIndex).flags.Comerciando = False
    Call WriteCommerceEnd(UserIndex)
    
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/03/2010
    '11/03/2010: ZaMa - Le avisa por consola al que cencela que dejo de comerciar.
    '***************************************************
    With UserList(UserIndex)
        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(.ComUsu.DestUsu)
                
            End If
            
        End If
        
        Call FinComerciarUsu(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has dejado de comerciar.", FontTypeNames.FONTTYPE_TALK)
        
    End With
    
End Sub

''
' Handles the "UserCommerceConfirm" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUserCommerceConfirm(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    
    'Validate the commerce
    If PuedeSeguirComerciando(UserIndex) Then
        'Tell the other user the confirmation of the offer
        Call WriteUserOfferConfirm(UserList(UserIndex).ComUsu.DestUsu)
        UserList(UserIndex).ComUsu.Confirmo = True
        
    End If
    
End Sub

Private Sub HandleCommerceChat(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 03/12/2009
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Chat    As String
        
        Chat = Reader.ReadString8()
        
        If LenB(Chat) <> 0 Then
            If PuedeSeguirComerciando(UserIndex) Then
                'Analize chat...
                Call Statistics.ParseChat(Chat)
                
                Chat = UserList(UserIndex).Name & "> " & Chat
                Call WriteCommerceChat(UserIndex, Chat, FontTypeNames.FONTTYPE_PARTY)
                Call WriteCommerceChat(UserList(UserIndex).ComUsu.DestUsu, Chat, FontTypeNames.FONTTYPE_PARTY)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(UserIndex)
        
    End With
    
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    'Trade accepted
    Call AceptarComercioUsu(UserIndex)
    
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim otherUser   As Integer
    
    With UserList(UserIndex)
        
        otherUser = .ComUsu.DestUsu
        
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)
                
            End If
            
        End If
        
        Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)
        
    End With
    
End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/25/09
    '07/25/09: Marco - Agregue un checkeo para patear a los usuarios que tiran items mientras comercian.
    '***************************************************
    
    Dim Slot        As Byte
    
    Dim Amount      As Integer
    
    With UserList(UserIndex)
        
        Slot = Reader.ReadInt8
        Amount = Reader.ReadInt16()
        
        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Muerto = 1 Or ((.flags.Privilegios And PlayerType.consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub
        
        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub
        
        'Si esta navegando y no es pirata, no dejamos tirar items al agua.
        If .flags.Navegando = 1 And Not .Clase = eClass.Pirat Then
            Call WriteConsoleMsg(UserIndex, "Solo los Piratas pueden tirar items en altamar", FontTypeNames.FONTTYPE_INFO)
            
            Exit Sub
        End If
        
        'Are we dropping gold or other items
        If Slot = FLAGORO Then
            If Amount > 10000 Then Exit Sub        'Don't drop too much gold
            
            Call TirarOro(Amount, UserIndex)
            
            Call WriteUpdateGold(UserIndex)
        Else
            
            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then
                    Exit Sub
                    
                End If
                
                Call DropObj(UserIndex, Slot, Amount, .Pos.Map, .Pos.X, .Pos.Y)
                
            End If
            
        End If
        
    End With
    
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '13/11/2009: ZaMa - Ahora los npcs pueden atacar al usuario si quizo castear un hechizo
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Spell   As Byte
        
        Spell = Reader.ReadInt8
        
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        If Spell < 1 Then
            .flags.Hechizo = 0
            Exit Sub
        ElseIf Spell > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
            Exit Sub
            
        End If
        
        .flags.Hechizo = .Stats.UserHechizos(Spell)
        
    End With
    
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim X           As Byte
    
    Dim Y           As Byte
    
    X = Reader.ReadInt8
    Y = Reader.ReadInt8
    
    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
    
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim X           As Byte
    
    Dim Y           As Byte
    
    X = Reader.ReadInt8
    Y = Reader.ReadInt8
    
    Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
    
End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010 (ZaMa)
    '13/01/2010: ZaMa - El pirata se puede ocultar en barca
    '***************************************************
    
    On Error GoTo errHandler
    With UserList(UserIndex)
        
        Dim Skill   As eSkill
        
1                        Skill = Reader.ReadInt8
        
2                        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
        'If exiting, cancel
3                        Call CancelExit(UserIndex)
        
        Select Case Skill
            
            Case Robar, Magia, Domar
4                                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, Skill)
                
            Case Ocultarse
                
                ' Verifico si se peude ocultar en este mapa
5                                                If (MapInfo(.Pos.Map).OcultarSinEfecto = 1) Or (MapInfo(.Pos.Map).InviSinEfecto = 1) Then
6                                                    Call WriteConsoleMsg(UserIndex, "Ocultarse no funciona aqui!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
            If .flags.EnConsulta Then
                Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si estas en consulta.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
            If .flags.Navegando = 1 Then
                If .Clase <> eClass.Pirat Then
                    
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 3 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si estas navegando.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 3
                        
                    End If
                    
                    '[/CDT]
                    Exit Sub
                    
                End If
                
            End If
            
            If .flags.Oculto = 1 Then
                
                '[CDT 17-02-2004]
                If Not .flags.UltimoMensaje = 2 Then
                    Call WriteConsoleMsg(UserIndex, "Ya estas oculto.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 2
                    
                End If
                
                '[/CDT]
                Exit Sub
                
            End If
            
7                                        Call DoOcultarse(UserIndex)
            
    End Select
    
End With
Exit Sub
errHandler:
Call LogError("Error en HandleWork en " & Erl & " - Skill: " & Skill & ". Err: " & Err.Number & " " & Err.description)
End Sub

''
' Handles the "InitCrafting" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInitCrafting(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/01/2010
    '
    '***************************************************
    Dim TotalItems    As Long
    
    Dim ItemsPorCiclo As Integer
    
    With UserList(UserIndex)
        
        TotalItems = Reader.ReadInt32
        ItemsPorCiclo = Reader.ReadInt16
        
        If TotalItems > 0 Then
            
            .Construir.Cantidad = TotalItems
            .Construir.PorCiclo = MinimoInt(MaxItemsConstruibles(UserIndex), ItemsPorCiclo)
            
        End If
        
    End With
    
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos.", FontTypeNames.FONTTYPE_FIGHT))
        Call WriteErrorMsg(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.")
        Call CloseSocket(UserIndex)
        
    End With
    
End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Slot    As Byte
        
        Slot = Reader.ReadInt8
        
        If Slot <= .CurrentInventorySlots And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
            
        End If
        
        If .flags.Meditando Then
            
            Exit Sub        'The error message should have been provided by the client.
            
        End If
        
        Call UseInvItem(UserIndex, Slot)
        
    End With
    
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Item        As Integer
    
    Item = Reader.ReadInt16()
    
    If Item < 1 Then Exit Sub
    
    If ObjData(Item).SkHerreria = 0 Then Exit Sub
    
    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
    Call HerreroConstruirItem(UserIndex, Item)
    
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    Dim Item        As Integer
    
1                    Item = Reader.ReadInt16()
    
2                    If Item < 1 Then Exit Sub
    
3                    If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
    
4                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
5                    Call CarpinteroConstruirItem(UserIndex, Item)
    
    Exit Sub
errHandler:
    Call LogError("Error en HandleCraftcarpenter en " & Erl & " - Item: " & Item & ". Err " & Err.Number & " " & Err.description)
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 14/01/2010 (ZaMa)
    '16/11/2009: ZaMa - Agregada la posibilidad de extraer madera elfica.
    '12/01/2010: ZaMa - Ahora se admiten armas arrojadizas (proyectiles sin municiones).
    '14/01/2010: ZaMa - Ya no se pierden municiones al atacar npcs con dueno.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim X           As Byte
        
        Dim Y           As Byte
        
        Dim Skill       As eSkill
        
        Dim DummyInt    As Integer
        
        Dim tU          As Integer        'Target user
        
        Dim tN          As Integer        'Target NPC
        
        Dim WeaponIndex As Integer
        
        X = Reader.ReadInt8
        Y = Reader.ReadInt8
        
        Skill = Reader.ReadInt8
        
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub
        
        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
            
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            
            Case eSkill.Proyectiles
                
                'Check attack interval
                If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub
                
                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Sub
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                Call LanzarProyectil(UserIndex, X, Y)
                
            Case eSkill.Magia
                
                'Check the map allows spells to be casted.
                If MapInfo(.Pos.Map).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energia.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .IP & " a la posicion (" & .Pos.Map & "/" & X & "/" & Y & ")")
                    Exit Sub
                End If
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                'Check Spell-Hit interval
                If Not IntervaloPermiteGolpeMagia(UserIndex) Then
                    
                    'Check Magic interval
                    If Not IntervaloPermiteLanzarSpell(UserIndex) Then
                        Exit Sub
                    End If
                    
                End If
                
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    .flags.Hechizo = 0
                Else
                    Call WriteConsoleMsg(UserIndex, "Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Case eSkill.pesca
                WeaponIndex = .Invent.WeaponEqpObjIndex
                
                If WeaponIndex = 0 Then Exit Sub
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.BAJOTECHO Or MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.CASA Then
                    Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                    
                End If
                
                If HayAgua(.Pos.Map, X, Y) Then
                    
                    Select Case WeaponIndex
                        
                        Case CANA_PESCA, CANA_PESCA_NEWBIE
                            Call DoPescar(UserIndex)
                            
                        Case RED_PESCA
                            
                            DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                            
                            If DummyInt = 0 Then
                                Call WriteConsoleMsg(UserIndex, "No hay un yacimiento de peces donde pescar.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                                
                            End If
                            
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                                
                            End If
                            
                            If .Pos.X = X And .Pos.Y = Y Then
                                Call WriteConsoleMsg(UserIndex, "No puedes pescar desde alli.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                                
                            End If
                            
                            'Hay un arbol normal donde clickeo
                            If ObjData(DummyInt).OBJType = eOBJType.otYacimientoPez Then
                                Call DoPescarRed(UserIndex)
                            Else
                                Call WriteConsoleMsg(UserIndex, "No hay un yacimiento de peces donde pescar.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                                
                            End If
                            
                        Case Else
                            
                            Exit Sub        'Invalid item!
                            
                    End Select
                    
                    'Play sound!
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Case eSkill.Robar
                
                'Does the map allow us to steal here
                If MapInfo(.Pos.Map).Pk Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> UserIndex Then
                        
                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub
                                    
                                End If
                                
                                '17/09/02
                                'Check the trigger
                                If MapData(UserList(tU).Pos.Map, X, Y).trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(UserIndex, "No puedes robar aqui.", FontTypeNames.FONTTYPE_WARNING)
                                    Exit Sub
                                    
                                End If
                                
                                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(UserIndex, "No puedes robar aqui.", FontTypeNames.FONTTYPE_WARNING)
                                    Exit Sub
                                    
                                End If
                                
                                Call DoRobar(UserIndex, tU)
                                
                            End If
                            
                        End If
                        
                    Else
                        Call WriteConsoleMsg(UserIndex, "No hay a quien robarle!", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Case eSkill.Talar
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                WeaponIndex = .Invent.WeaponEqpObjIndex
                
                If WeaponIndex = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Deberias equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                    
                End If
                
                If WeaponIndex <> HACHA_LENADOR And WeaponIndex <> HACHA_LENA_ELFICA And WeaponIndex <> HACHA_LENADOR_NEWBIE Then
                    ' Podemos llegar aca si el user equipo el anillo dsp de la U y antes del click
                    Exit Sub
                    
                End If
                
                DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                If DummyInt > 0 Then
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                        
                    End If
                    
                    'Barrin 29/9/03
                    If .Pos.X = X And .Pos.Y = Y Then
                        Call WriteConsoleMsg(UserIndex, "No puedes talar desde alli.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                        
                    End If
                    
                    'Hay un arbol normal donde clickeo
                    If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                        If WeaponIndex = HACHA_LENADOR Or WeaponIndex = HACHA_LENADOR_NEWBIE Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                            Call DoTalar(UserIndex)
                        Else
                            Call WriteConsoleMsg(UserIndex, "No puedes extraer lena de este arbol con este hacha.", FontTypeNames.FONTTYPE_INFO)
                            
                        End If
                        
                        ' Arbol Elfico
                    ElseIf ObjData(DummyInt).OBJType = eOBJType.otArbolElfico Then
                        
                        If WeaponIndex = HACHA_LENA_ELFICA Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                            Call DoTalar(UserIndex, True)
                        Else
                            Call WriteConsoleMsg(UserIndex, "El hacha utilizado no es suficientemente poderosa.", FontTypeNames.FONTTYPE_INFO)
                            
                        End If
                        
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay ningUn arbol ahi.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Case eSkill.Mineria
                
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                WeaponIndex = .Invent.WeaponEqpObjIndex
                
                If WeaponIndex = 0 Then Exit Sub
                
                If WeaponIndex <> PIQUETE_MINERO And WeaponIndex <> PIQUETE_MINERO_NEWBIE Then
                    ' Podemos llegar aca si el user equipo el anillo dsp de la U y antes del click
                    Exit Sub
                    
                End If
                
                'Target whatever is in the tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                If DummyInt > 0 Then
                    
                    'Check distance
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                        
                    End If
                    
                    'Hay un yacimiento donde clickeo
                    If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
                        Call DoMineria(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahi no hay ningUn yacimiento.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                    'Fix Sonido by Fakkerz 24/05/2020
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_MINERO, .Pos.X, .Pos.Y))
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahi no hay ningUn yacimiento.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                
                'Target whatever is that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                tN = .flags.TargetNPC
                
                If tN > 0 Then
                    If Npclist(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                            
                        End If
                        
                        If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes domar una criatura que esta luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                            
                        End If
                        
                        Call DoDomar(UserIndex, tN)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay ninguna criatura alli!", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Case FundirMetal        'UGLY!!! This is a constant, not a skill!!
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                'Check there is a proper item there
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
                        
                        'Validate other items
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > .CurrentInventorySlots Then
                            Exit Sub
                            
                        End If
                        
                        ''chequeamos que no se zarpe duplicando oro
                        If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                            If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                                Call WriteConsoleMsg(UserIndex, "No tienes mas minerales.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                                
                            End If
                            
                            ''FUISTE
                            Call WriteErrorMsg(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                            Call CloseSocket(UserIndex)
                            Exit Sub
                            
                        End If
                        
                        If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales Then
                            Call FundirMineral(UserIndex)
                        ElseIf ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otWeapon Then
                            Call FundirArmas(UserIndex)
                            
                        End If
                        
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahi no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahi no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Case eSkill.Herreria
                'Target wehatever is in that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(UserIndex)
                        Call EnivarArmadurasConstruibles(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahi no hay ningUn yunque.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "Ahi no hay ningUn yunque.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
        End Select
        
    End With
    
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/11/09
    '05/11/09: Pato - Ahora se quitan los espacios del principio y del fin del nombre del clan
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Desc      As String
        
        Dim GuildName As String
        
        Dim Site      As String
        
        Dim codex()   As String
        
        Dim errorStr  As String
        
        Desc = Reader.ReadString8()
        GuildName = Trim$(Reader.ReadString8())
        Site = Reader.ReadString8()
        codex = Split(Reader.ReadString8(), SEPARATOR)
        
        If modGuilds.CrearNuevoClan(UserIndex, Desc, GuildName, Site, codex, .FundandoGuildAlineacion, errorStr) Then
            Dim Message As String
            Message = .Name & " fundo el clan " & GuildName & " de alineacion " & modGuilds.GuildAlignment(.GuildIndex)
            
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
            
            'Update tag
            Call RefreshCharStatus(UserIndex)
            
            'Aqui solo vamos a hacer un request a los endpoints de la aplicacion en Node.js
            'el repositorio para hacer funcionar esto, es este: https://github.com/ao-libre/ao-api-server
            'Si no tienen interes en usarlo pueden desactivarlo en el Server.ini
            If ConexionAPI Then
                Call ApiEndpointSendNewGuildCreatedMessageDiscord(Message, Desc, GuildName, Site)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim itemSlot As Byte
        
        itemSlot = Reader.ReadInt8
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then Exit Sub
        
        'Validate item slot
        If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
        Call EquiparInvItem(UserIndex, itemSlot)
        
    End With
    
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 06/28/2008
    'Last Modified By: NicoNZ
    ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
    ' 06/28/2008: NicoNZ - Solo se puede cambiar si esta inmovilizado.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim heading As eHeading
        
        Dim posX    As Integer
        
        Dim posY    As Integer
        
        heading = Reader.ReadInt8
        
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
            
            If LegalPos(.Pos.Map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                Exit Sub
                
            End If
            
        End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If heading > 0 And heading < 5 Then
            .Char.heading = heading
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
            
        End If
        
    End With
    
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/19/09
    '11/19/09: Pato - Adapting to new skills system.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim i                      As Long
        
        Dim Count                  As Integer
        
        Dim points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        For i = 1 To NUMSKILLS
            points(i) = Reader.ReadInt8
            
            If points(i) < 0 Then
                Call LogHackAttemp(.Name & " IP:" & .IP & " trato de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub
                
            End If
            
            Count = Count + points(i)
        Next i
        
        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.Name & " IP:" & .IP & " trato de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
            
        End If
        
        .Counters.AsignedSkills = MinimoInt(10, .Counters.AsignedSkills + Count)
        
        With .Stats
            
            For i = 1 To NUMSKILLS
                
                If points(i) > 0 Then
                    .SkillPts = .SkillPts - points(i)
                    .UserSkills(i) = .UserSkills(i) + points(i)
                    
                    'Client should prevent this, but just in case...
                    If .UserSkills(i) > 100 Then
                        .SkillPts = .SkillPts + .UserSkills(i) - 100
                        .UserSkills(i) = 100
                        
                    End If
                    
                    Call CheckEluSkill(UserIndex, i, True)
                    
                End If
                
            Next i
            
        End With
        
    End With
    
End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim SpawnedNpc As Integer
        
        Dim PetIndex   As Byte
        
        PetIndex = Reader.ReadInt8
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If PetIndex > 0 And PetIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(PetIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1
                    
                End If
                
            End If
            
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer mas criaturas, mata las existentes.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            
        End If
        
    End With
    
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Slot    As Byte
        
        Dim Amount  As Integer
        
        Slot = Reader.ReadInt8
        Amount = Reader.ReadInt16()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'El target es un NPC valido
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'El NPC puede comerciar
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningun interes en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
            
        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No estas comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, Slot, Amount)
        
    End With
    
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Slot    As Byte
        
        Dim Amount  As Integer
        
        Slot = Reader.ReadInt8
        Amount = Reader.ReadInt16()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'El target es un NPC valido
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'Es el banquero
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
            
        End If
        
        'User retira el item del slot
        Call UserRetiraItem(UserIndex, Slot, Amount)
        
    End With
    
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Slot    As Byte
        
        Dim Amount  As Integer
        
        Slot = Reader.ReadInt8
        Amount = Reader.ReadInt16()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'El target es un NPC valido
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'El NPC puede comerciar
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningun interes en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
            
        End If
        
        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, Slot, Amount)
        
    End With
    
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Slot    As Byte
        
        Dim Amount  As Integer
        
        Slot = Reader.ReadInt8
        Amount = Reader.ReadInt16()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'El target es un NPC valido
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'El NPC puede comerciar
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
            
        End If
        
        'User deposita el item del slot rdata
        Call UserDepositaItem(UserIndex, Slot, Amount)
        
    End With
    
End Sub

''
' Handles the "ForumPost" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForumPost(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 02/01/2010
    '02/01/2010: ZaMa - Implemento nuevo sistema de foros
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim ForumMsgType As eForumMsgType
        
        Dim File         As String
        
        Dim Title        As String
        
        Dim Post         As String
        
        Dim ForumIndex   As Integer
        
        Dim postFile     As String
        
        Dim ForumType    As Byte
        
        ForumMsgType = Reader.ReadInt8
        
        Title = Reader.ReadString8()
        Post = Reader.ReadString8()
        
        If .flags.TargetObj > 0 Then
            ForumType = ForumAlignment(ForumMsgType)
            
            Select Case ForumType
                
                Case eForumType.ieGeneral
                    ForumIndex = GetForumIndex(ObjData(.flags.TargetObj).ForoID)
                    
                Case eForumType.ieREAL
                    ForumIndex = GetForumIndex(FORO_REAL_ID)
                    
                Case eForumType.ieCAOS
                    ForumIndex = GetForumIndex(FORO_CAOS_ID)
                    
            End Select
            
            Call AddPost(ForumIndex, Post, .Name, Title, EsAnuncio(ForumMsgType))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Dir         As Integer
    
    If Reader.ReadBool() Then
        Dir = 1
    Else
        Dir = -1
        
    End If
    
    Call DesplazarHechizo(UserIndex, Dir, Reader.ReadInt8)
    
End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 06/14/09
    '
    '***************************************************
    
    Dim Dir         As Integer
    
    Dim Slot        As Byte
    
    Dim TempItem    As obj
    
    If Reader.ReadBool() Then
        Dir = 1
    Else
        Dir = -1
        
    End If
    
    Slot = Reader.ReadInt8
    
    With UserList(UserIndex)
        TempItem.ObjIndex = .BancoInvent.Object(Slot).ObjIndex
        TempItem.Amount = .BancoInvent.Object(Slot).Amount
        
        If Dir = 1 Then        'Mover arriba
        .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
        .BancoInvent.Object(Slot - 1).ObjIndex = TempItem.ObjIndex
        .BancoInvent.Object(Slot - 1).Amount = TempItem.Amount
        
        Call UpdateBanUserInv(False, UserIndex, Slot - 1)
    Else        'mover abajo
        .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
        .BancoInvent.Object(Slot + 1).ObjIndex = TempItem.ObjIndex
        .BancoInvent.Object(Slot + 1).Amount = TempItem.Amount
        
        Call UpdateBanUserInv(False, UserIndex, Slot + 1)
    End If
    
    Call UpdateBanUserInv(False, UserIndex, Slot)
    
End With

End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Desc    As String
        
        Dim codex() As String
        
        Desc = Reader.ReadString8()
        codex = Split(Reader.ReadString8(), SEPARATOR)
        
        Call modGuilds.ChangeCodexAndDesc(Desc, codex, .GuildIndex)
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 24/11/2009
    '24/11/2009: ZaMa - Nuevo sistema de comercio
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Amount    As Long
        
        Dim Slot      As Byte
        
        Dim tUser     As Integer
        
        Dim OfferSlot As Byte
        
        Dim ObjIndex  As Integer
        
        Slot = Reader.ReadInt8
        Amount = Reader.ReadInt32()
        OfferSlot = Reader.ReadInt8
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        ' If he's already confirmed his offer, but now tries to change it, then he's cheating
        If UserList(UserIndex).ComUsu.Confirmo = True Then
            
            ' Finish the trade
            Call FinComerciarUsu(UserIndex)
            
            If tUser <= 0 Or tUser > MaxUsers Then
                Call FinComerciarUsu(tUser)
                
            End If
            
            Exit Sub
            
        End If
        
        'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
        If ((Slot < 0 Or Slot > UserList(UserIndex).CurrentInventorySlots) And Slot <> FLAGORO) Then Exit Sub
        
        'If OfferSlot is invalid, then ignore it.
        If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 1 Then Exit Sub
        
        ' Can be negative if substracted from the offer, but never 0.
        If Amount = 0 Then Exit Sub
        
        'Has he got enough
        If Slot = FLAGORO Then
            
            ' Can't offer more than he has
            If Amount > .Stats.Gld - .ComUsu.GoldAmount Then
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
                
            End If
            
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.GoldAmount Then
                    Amount = .ComUsu.GoldAmount * (-1)
                    
                End If
                
            End If
            
        Else
            
            'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
            If Slot <> 0 Then ObjIndex = .Invent.Object(Slot).ObjIndex
            
            ' Can't offer more than he has
            If Not HasEnoughItems(UserIndex, ObjIndex, TotalOfferItems(ObjIndex, UserIndex) + Amount) Then
                
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
                
            End If
            
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.cant(OfferSlot) Then
                    Amount = .ComUsu.cant(OfferSlot) * (-1)
                    
                End If
                
            End If
            
            If ItemNewbie(ObjIndex) Then
                Call WriteCancelOfferItem(UserIndex, OfferSlot)
                Exit Sub
                
            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteCommerceChat(UserIndex, "No puedes vender tu barco mientras lo estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                    
                End If
                
            End If
            
            If .flags.Equitando = 1 Then
                If .Invent.MonturaEqpSlot = Slot Then
                    Call WriteConsoleMsg(UserIndex, "No podes vender tu montura mientras lo estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            
            If .Invent.MochilaEqpSlot > 0 Then
                If .Invent.MochilaEqpSlot = Slot Then
                    Call WriteCommerceChat(UserIndex, "No puedes vender tu alforja o mochila mientras la estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                    
                End If
                
            End If
            
        End If
        
        Call AgregarOferta(UserIndex, OfferSlot, ObjIndex, Amount, Slot = FLAGORO)
        Call EnviarOferta(tUser, OfferSlot)
        
    End With
    
End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Guild          As String
        
        Dim errorStr       As String
        
        Dim otherClanIndex As String
        
        Guild = Reader.ReadString8()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(UserIndex, Guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Guild          As String
        
        Dim errorStr       As String
        
        Dim otherClanIndex As String
        
        Guild = Reader.ReadString8()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, Guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & Guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Guild          As String
        
        Dim errorStr       As String
        
        Dim otherClanIndex As String
        
        Guild = Reader.ReadString8()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(UserIndex, Guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Guild          As String
        
        Dim errorStr       As String
        
        Dim otherClanIndex As String
        
        Guild = Reader.ReadString8()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, Guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Guild    As String
        
        Dim proposal As String
        
        Dim errorStr As String
        
        Guild = Reader.ReadString8()
        proposal = Reader.ReadString8()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, Guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de paz enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Guild    As String
        
        Dim proposal As String
        
        Dim errorStr As String
        
        Guild = Reader.ReadString8()
        proposal = Reader.ReadString8()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, Guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de alianza enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Guild    As String
        
        Dim errorStr As String
        
        Dim details  As String
        
        Guild = Reader.ReadString8()
        
        details = modGuilds.r_VerPropuesta(UserIndex, Guild, RELACIONES_GUILD.ALIADOS, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Guild    As String
        
        Dim errorStr As String
        
        Dim details  As String
        
        Guild = Reader.ReadString8()
        
        details = modGuilds.r_VerPropuesta(UserIndex, Guild, RELACIONES_GUILD.PAZ, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim User    As String
        
        Dim details As String
        
        User = Reader.ReadString8()
        
        details = modGuilds.a_DetallesAspirante(UserIndex, User)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje no ha mandado solicitud, o no estas habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(UserIndex, details)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call WriteAlianceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.ALIADOS))
    
End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call WritePeaceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.PAZ))
    
End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Guild           As String
        
        Dim errorStr        As String
        
        Dim otherGuildIndex As Integer
        
        Guild = Reader.ReadString8()
        
        otherGuildIndex = modGuilds.r_DeclararGuerra(UserIndex, Guild, errorStr)
        
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Call modGuilds.ActualizarWebSite(UserIndex, Reader.ReadString8())
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim errorStr As String
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If Not modGuilds.a_AceptarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
                Call RefreshCharStatus(tUser)
                
            End If
            
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim errorStr As String
        
        Dim UserName As String
        
        Dim Reason   As String
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        Reason = Reader.ReadString8()
        
        If Not modGuilds.a_RechazarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, errorStr & " : " & Reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, Reason)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName   As String
        
        Dim GuildIndex As Integer
        
        UserName = Reader.ReadString8()
        
        GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
        
        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            Call WriteConsoleMsg(UserIndex, "No puedes expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Call modGuilds.ActualizarNoticias(UserIndex, Reader.ReadString8())
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Call modGuilds.SendDetallesPersonaje(UserIndex, Reader.ReadString8())
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        Dim Error   As String
        
        If Not modGuilds.v_AbrirElecciones(UserIndex, Error) Then
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .Name, FontTypeNames.FONTTYPE_GUILD))
            
        End If
        
    End With
    
End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Guild       As String
        
        Dim application As String
        
        Dim errorStr    As String
        
        Guild = Reader.ReadString8()
        application = Reader.ReadString8()
        
        If Not modGuilds.a_NuevoAspirante(UserIndex, Guild, application, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del lider de " & Guild & ".", FontTypeNames.FONTTYPE_GUILD)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Call modGuilds.SendGuildDetails(UserIndex, Reader.ReadString8())
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

Private Sub WriteConsoleServerUpTimeMsg(ByVal UserIndex As Integer)
    Dim Time        As Long
    Dim UpTimeStr   As String
    
    'Get total time in seconds
    Time = ((GetTickCount()) - tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (Time Mod 60) & " segundos."
    Time = Time \ 60
    
    UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
    Time = Time \ 60
    
    UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
    Time = Time \ 24
    
    If Time = 1 Then
        UpTimeStr = Time & " dia, " & UpTimeStr
    Else
        UpTimeStr = Time & " dias, " & UpTimeStr
    End If
    
    Call WriteConsoleMsg(UserIndex, "Tiempo del Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub

''
' Handles the "Online" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 03/07/2020 (Jopi)
    '14/07/19 (Recox) - Ahora se muestra una lista de nombres de jugadores online, se suman los gms tambien a la lista.
    '03/07/2020 (Jopi) - Usamos la clase cStringBuilder para generar la lista de usuarios online.
    '03/07/2020 (Jopi) - Mostramos quienes son trabajadores y cuantos estan online.
    '***************************************************
    
    ' Generador de strings
    Dim SB          As cStringBuilder
    Set SB = New cStringBuilder
    
    ' Contadores
    Dim i           As Long
    Dim Count       As Long
    Dim CountTrabajadores As Long
    
    With UserList(UserIndex)
        
        ' Recorremos la lista de usuarios online.
        For i = 1 To LastUser
            
            If LenB(UserList(i).Name) <> 0 Then
                
                ' Agregamos el nombre al final del string.
                Call SB.Append(UserList(i).Name)
                
                ' Si es trabajador...
                If UserList(i).Clase = eClass.Worker Then
                    
                    ' Si es Cazador y tiene 100 en supervivencia o es Game Master.
                    If EsGm(UserIndex) Or (.Clase = eClass.Hunter And .Stats.UserSkills(eSkill.Supervivencia) = 100) Then
                        ' Incrementamos en 1 el contador de jugadores online que son trabajadores.
                        CountTrabajadores = CountTrabajadores + 1
                        ' Agregamos un sufijo que indique que es un trabajador.
                        Call SB.Append(" [T]")
                    End If
                    
                End If
                
                ' Si no terminamos de recorrer la lista, agregamos el separador.
                If i <> LastUser Then
                    Call SB.Append(", ")
                End If
                
                ' Incrementa en 1 el contador de jugadores online TOTALES
                Count = Count + 1
                
            End If
            
        Next i
        
        Call WriteConsoleMsg(UserIndex, SB.toString, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Usuarios en linea: " & CStr(Count), FontTypeNames.FONTTYPE_INFOBOLD)
        
        ' Si es Cazador y tiene 100 en supervivencia o es Game Master.
        If EsGm(UserIndex) Or (.Clase = eClass.Hunter And .Stats.UserSkills(eSkill.Supervivencia) = 100) Then
            Call WriteConsoleMsg(UserIndex, "Trabajadores en linea:" & CStr(CountTrabajadores), FontTypeNames.FONTTYPE_INFOBOLD)
        End If
        
        ' Liberamos los recursos del generador de strings
        Set SB = Nothing
        
    End With
    
    Call WriteConsoleServerUpTimeMsg(UserIndex)
    
End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ)
    'If user is invisible, it automatically becomes
    'visible before doing the countdown to exit
    '15/04/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
    '13/01/2020 - Se pusieron nuevas validaciones para las monturas. (Recox)
    '***************************************************
    Dim tUser        As Integer
    
    With UserList(UserIndex)
        
        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
            
        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_WARNING)
                    Call FinComerciarUsu(tUser)
                    
                End If
                
            End If
            
            Call WriteConsoleMsg(UserIndex, "Comercio cancelado.", FontTypeNames.FONTTYPE_WARNING)
            Call FinComerciarUsu(UserIndex)
            
        End If
        
        Call Cerrar_Usuario(UserIndex)
        
    End With
    
End Sub

''
' Handles the "GuildLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim GuildIndex  As Integer
    
    With UserList(UserIndex)
        
        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(UserIndex, .Name)
        
        If GuildIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(UserIndex, "Tu no puedes salir de este clan.", FontTypeNames.FONTTYPE_GUILD)
            
        End If
        
    End With
    
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim earnings    As Integer
    
    Dim Percentage  As Integer
    
    With UserList(UserIndex)
        
        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        Select Case Npclist(.flags.TargetNPC).NPCtype
            
            Case eNPCType.Banquero
                Call WriteChatOverHead(UserIndex, "Tienes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
            Case eNPCType.Timbero
                
                If Not .flags.Privilegios And PlayerType.User Then
                    earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
                    If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                        Percentage = Int(earnings * 100 / Apuestas.Ganancias)
                        
                    End If
                    
                    If earnings < 0 And Apuestas.Perdidas <> 0 Then
                        Percentage = Int(earnings * 100 / Apuestas.Perdidas)
                        
                    End If
                    
                    Call WriteConsoleMsg(UserIndex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & Percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
        End Select
        
    End With
    
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'Make sure it's his pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it!
        Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO
        
        Call Expresar(.flags.TargetNPC, UserIndex)
        
    End With
    
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, UserIndex)
        
    End With
    
End Sub

''
' Handles the "ReleasePet" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReleasePet(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/11/2009
    '
    '***************************************************
    With UserList(UserIndex)
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar una mascota, haz click izquierdo sobre ella.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'Do it
        Call QuitarPet(UserIndex, .flags.TargetNPC)
        
    End With
    
End Sub

''
' Handles the "TrainList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'Make sure it's the trainer
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)
        
    End With
    
End Sub

''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(UserIndex, "Te acomodas junto a la fogata y comienzas a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
            .flags.Descansar = Not .flags.Descansar
        Else
            
            If .flags.Descansar Then
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
                .flags.Descansar = False
                Exit Sub
                
            End If
            
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/08 (NicoNZ)
    'Arregle un bug que mandaba un index de la meditacion diferente
    'al que decia el server.
    '***************************************************
    With UserList(UserIndex)
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes meditar cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        If .flags.Equitando Then
            Call WriteConsoleMsg(UserIndex, "No puedes meditar mientras si estas montado.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Can he meditate
        If .Stats.MaxMAN = 0 Then
            Call WriteConsoleMsg(UserIndex, "Solo las clases magicas conocen el arte de la meditacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'Admins don't have to wait :D
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMAN = .Stats.MaxMAN
            Call WriteConsoleMsg(UserIndex, "Mana restaurado.", FontTypeNames.FONTTYPE_VENENO)
            Call WriteUpdateMana(UserIndex)
            Exit Sub
            
        End If
        
        Call WriteMeditateToggle(UserIndex)
        
        If .flags.Meditando Then Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        
        .flags.Meditando = Not .flags.Meditando
        
        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            .Counters.tInicioMeditar = GetTickCount()
            
            Call WriteConsoleMsg(UserIndex, "Te estas concentrando. En " & Fix(TIEMPO_INICIOMEDITAR / 1000) & " segundos comenzaras a meditar.", FontTypeNames.FONTTYPE_INFO)
            
            .Char.loops = INFINITE_LOOPS
            
            'Show proper FX according to level
            If .Stats.ELV < 13 Then
                .Char.FX = FXIDs.FXMEDITARCHICO
                
            ElseIf .Stats.ELV < 25 Then
                .Char.FX = FXIDs.FXMEDITARMEDIANO
                
            ElseIf .Stats.ELV < 35 Then
                .Char.FX = FXIDs.FXMEDITARGRANDE
                
            ElseIf .Stats.ELV < 42 Then
                .Char.FX = FXIDs.FXMEDITARXGRANDE
                
            Else
                .Char.FX = FXIDs.FXMEDITARXXGRANDE
                
            End If
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
        Else
            .Counters.bPuedeMeditar = False
            
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            
        End If
        
    End With
    
End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/01/20
    'Arreglo validacion de NPC para que funcione el comando. (Recox)
    '***************************************************
    With UserList(UserIndex)
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'Validate NPC and make sure player is dead
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 5 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        Call SacerdoteResucitateUser(UserIndex)
    End With
    
End Sub

''
' Handles the "Consultation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConsultation(ByVal UserIndex As String)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 01/05/2010
    'Habilita/Deshabilita el modo consulta.
    '01/05/2010: ZaMa - Agrego validaciones.
    '16/09/2010: ZaMa - No se hace visible en los clientes si estaba navegando (porque ya lo estaba).
    '***************************************************
    
    Dim UserConsulta As Integer
    
    With UserList(UserIndex)
        
        ' Comando exclusivo para gms
        If Not EsGm(UserIndex) Then Exit Sub
        
        UserConsulta = .flags.TargetUser
        
        'Se asegura que el target es un usuario
        If UserConsulta = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        ' No podes ponerte a vos mismo en modo consulta.
        If UserConsulta = UserIndex Then Exit Sub
        
        ' No podes estra en consulta con otro gm
        If EsGm(UserConsulta) Then
            Call WriteConsoleMsg(UserIndex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        Dim UserName As String
        
        UserName = UserList(UserConsulta).Name
        
        ' Si ya estaba en consulta, termina la consulta
        If UserList(UserConsulta).flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Termino consulta con " & UserName)
            
            UserList(UserConsulta).flags.EnConsulta = False
            
            ' Sino la inicia
        Else
            Call WriteConsoleMsg(UserIndex, "Has iniciado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Inicio consulta con " & UserName)
            
            With UserList(UserConsulta)
                .flags.EnConsulta = True
                
                ' Pierde invi u ocu
                If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    .flags.invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    
                    If UserList(UserConsulta).flags.Navegando = 0 Then
                        Call UsUaRiOs.SetInvisible(UserConsulta, UserList(UserConsulta).Char.CharIndex, False)
                        
                    End If
                    
                End If
                
            End With
            
        End If
        
        Call UsUaRiOs.SetConsulatMode(UserConsulta)
        
    End With
    
End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor) Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        Call SacerdoteHealUser(UserIndex)
    End With
    
End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call SendUserStatsTxt(UserIndex, UserIndex)
    
End Sub

''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call SendHelp(UserIndex)
    
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i           As Integer
    
    With UserList(UserIndex)
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'Is it already in commerce mode
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estas comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            
            'Does the NPC want to trade
            If Npclist(.flags.TargetNPC).Comercia = 0 Then
                
                If LenB(Npclist(.flags.TargetNPC).Desc) <> 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ningun interes en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                End If
                
                Exit Sub
                
            End If
            
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
            '[Alejo]
        ElseIf .flags.TargetUser > 0 Then
            
            'User commerce...
            'Can he commerce
            If .flags.Privilegios And PlayerType.consejero Then
                Call WriteConsoleMsg(UserIndex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
                
            End If
            
            'Is the other one dead
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
            'Is it me
            If .flags.TargetUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
            'Is he already trading?? is it with me or someone else
            If UserList(.flags.TargetUser).flags.Comerciando = True And UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).Name
            
            For i = 1 To MAX_OFFER_SLOTS
                .ComUsu.cant(i) = 0
                .ComUsu.Objeto(i) = 0
            Next i
            
            .ComUsu.GoldAmount = 0
            
            .ComUsu.Acepto = False
            .ComUsu.Confirmo = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estas comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
            'If it's the banker....
            If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                Call IniciarDeposito(UserIndex)
                
            End If
            
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte mas.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(UserIndex)
        Else
            Call EnlistarCaos(UserIndex)
            
        End If
        
    End With
    
End Sub

''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim Matados     As Integer
    
    Dim NextRecom   As Integer
    
    Dim Diferencia  As Integer
    
    With UserList(UserIndex)
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        NextRecom = .Faccion.NextRecompensa
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
                
            End If
            
            Matados = .Faccion.CriminalesMatados
            Diferencia = NextRecom - Matados
            
            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, mata " & Diferencia & " criminales mas y te dare una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, y ya has matado los suficientes como para merecerte una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
            End If
            
        Else
            
            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a la legion oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
                
            End If
            
            Matados = .Faccion.CiudadanosMatados
            Diferencia = NextRecom - Matados
            
            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, mata " & Diferencia & " ciudadanos mas y te dare una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, y creo que estas en condiciones de merecer una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
            End If
            
        End If
        
    End With
    
End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
                
            End If
            
            Call RecompensaArmadaReal(UserIndex)
        Else
            
            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a la legion oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub
                
            End If
            
            Call RecompensaCaos(UserIndex)
            
        End If
        
    End With
    
End Sub

''
' Handles the "RequestMOTD" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMOTD(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call SendMOTD(UserIndex)
    
End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/10/08
    '01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
    '***************************************************
    
    Dim Time        As Long
    
    Dim UpTimeStr   As String
    
    Call WriteConsoleServerUpTimeMsg(UserIndex)
End Sub

''
' Handles the "PartyLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyLeave(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call mdParty.SalirDeParty(UserIndex)
    
End Sub

''
' Handles the "PartyCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyCreate(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
    
    Call mdParty.CrearParty(UserIndex)
    
End Sub

''
' Handles the "PartyJoin" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyJoin(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call mdParty.SolicitarIngresoAParty(UserIndex)
    
End Sub

''
' Handles the "ShareNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShareNpc(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/04/2010
    'Shares owned npcs with other user
    '***************************************************
    
    Dim TargetUserIndex  As Integer
    
    Dim SharingUserIndex As Integer
    
    With UserList(UserIndex)
        
        ' Didn't target any user
        TargetUserIndex = .flags.TargetUser
        
        If TargetUserIndex = 0 Then Exit Sub
        
        ' Can't share with admins
        If EsGm(TargetUserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con administradores!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        ' Pk or Caos
        If criminal(UserIndex) Then
            
            ' Caos can only share with other caos
            If esCaos(UserIndex) Then
                If Not esCaos(TargetUserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Solo puedes compartir npcs con miembros de tu misma faccion!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                    
                End If
                
                ' Pks don't need to share with anyone
            Else
                Exit Sub
                
            End If
            
            ' Ciuda or Army
        Else
            
            ' Can't share
            If criminal(TargetUserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con criminales!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
        End If
        
        ' Already sharing with target
        SharingUserIndex = .flags.ShareNpcWith
        
        If SharingUserIndex = TargetUserIndex Then Exit Sub
        
        ' Aviso al usuario anterior que dejo de compartir
        If SharingUserIndex <> 0 Then
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
        .flags.ShareNpcWith = TargetUserIndex
        
        Call WriteConsoleMsg(TargetUserIndex, .Name & " ahora comparte sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Ahora compartes tus npcs con " & UserList(TargetUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handles the "StopSharingNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleStopSharingNpc(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/04/2010
    'Stop Sharing owned npcs with other user
    '***************************************************
    
    Dim SharingUserIndex As Integer
    
    With UserList(UserIndex)
        
        SharingUserIndex = .flags.ShareNpcWith
        
        If SharingUserIndex <> 0 Then
            
            ' Aviso al que compartia y al que le compartia.
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SharingUserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
            
            .flags.ShareNpcWith = 0
            
        End If
        
    End With
    
End Sub

''
' Handles the "Inquiry" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    ConsultaPopular.SendInfoEncuesta (UserIndex)
    
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 15/07/2009
    '02/03/2009: ZaMa - Arreglado un indice mal pasado a la funcion de cartel de clanes overhead.
    '15/07/2009: ZaMa - Now invisible admins only speak by console
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Chat    As String
        
        Chat = Reader.ReadString8()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .GuildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & Chat))
                
                If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageChatOverHead("< " & Chat & " >", .Char.CharIndex, vbYellow))
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "PartyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyMessage(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Chat    As String
        
        Chat = Reader.ReadString8()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            Call mdParty.BroadCastParty(UserIndex, Chat)
            
            'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
            'Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).Pos.map, "||" & vbYellow & "Â°< " & mid$(rData, 7) & " >Â°" & CStr(UserList(UserIndex).Char.CharIndex))
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "CentinelReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCentinelReport(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 02/05/2012
    '                         Nuevo centinela (maTih.-)
    '***************************************************
    
    Call modCentinela.IngresaClave(UserIndex, Reader.ReadString8())
    
End Sub

''
' Handles the "GuildOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        Dim onlineList As String
        
        onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .GuildIndex)
        
        If .GuildIndex <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Companeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            Call WriteConsoleMsg(UserIndex, "No pertences a ningUn clan.", FontTypeNames.FONTTYPE_GUILDMSG)
            
        End If
        
    End With
    
End Sub

''
' Handles the "PartyOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyOnline(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Call mdParty.OnlineParty(UserIndex)
    
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Chat    As String
        
        Chat = Reader.ReadString8()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim request As String
        
        request = Reader.ReadString8()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.Name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If Not Ayuda.Existe(.Name) Then
            Call WriteConsoleMsg(UserIndex, "El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM.", FontTypeNames.FONTTYPE_INFO)
            Call Ayuda.Push(.Name)
        Else
            Call Ayuda.Quitar(.Name)
            Call Ayuda.Push(.Name)
            Call WriteConsoleMsg(UserIndex, "Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name + " ha solicitado la ayuda de algun GM con /GM. Podes usar el comando /SHOW SOS para ver quienes necesitan ayuda", FontTypeNames.FONTTYPE_INFO))
    End With
    
End Sub

''
' Handles the "BugReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBugReport(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim bugReport As String
        Dim n       As Integer
        
        bugReport = Reader.ReadString8()
        
        n = FreeFile
        Open App.Path & "\LOGS\BUGs.log" For Append Shared As n
        Print #n, "Usuario:" & .Name & "  Fecha:" & Date & "    Hora:" & Time
        Print #n, "BUG:"
        Print #n, bugReport
        Print #n, "########################################################################"
        Close #n
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim description As String
        
        description = Reader.ReadString8()
        
        If Not AsciiValidos(description) Then
            Call WriteConsoleMsg(UserIndex, "La descripcion tiene caracteres invalidos.", FontTypeNames.FONTTYPE_INFO)
        Else
            .Desc = Trim$(description)
            Call WriteConsoleMsg(UserIndex, "La descripcion ha cambiado.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim vote     As String
        
        Dim errorStr As String
        
        vote = Reader.ReadString8()
        
        If Not modGuilds.v_UsuarioVota(UserIndex, vote, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ShowGuildNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowGuildNews(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMA
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Call modGuilds.SendGuildNews(UserIndex)
        
    End With
    
End Sub

''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 25/08/2009
    '25/08/2009: ZaMa - Now only admins can see other admins' punishment list
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Name    As String
        
        Dim Count   As Integer
        
        Name = Reader.ReadString8()
        
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
            
            If (EsAdmin(Name) Or EsDios(Name) Or EsSemiDios(Name) Or EsConsejero(Name) Or EsRolesMaster(Name)) And (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
                Call WriteConsoleMsg(UserIndex, "No puedes ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
            Else
                
                If PersonajeExiste(Name) Then
                    Count = GetUserAmountOfPunishments(Name)
                    
                    If Count = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call SendUserPunishments(UserIndex, Name, Count)
                        
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ChangePassword" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangePassword(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Creation Date: 10/10/07
    'Last Modified By: Rapsodius
    '***************************************************
    
    'SHA256
    Dim oSHA256     As CSHA256
    
    Set oSHA256 = New CSHA256
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim oldSalt    As String
        
        Dim Salt       As String
        
        Dim oldPass    As String
        
        Dim newPass    As String
        
        Dim storedPass As String
        
        'Hasheamos el pass junto al Salt
        oldSalt = GetUserSalt(UserList(UserIndex).Name)
        oldPass = oSHA256.SHA256(Reader.ReadString8() & oldSalt)
        
        'Asignamos un nuevo Salt y lo hasheamos junto al nuevo pass
        Salt = RandomString(10)
        newPass = oSHA256.SHA256(Reader.ReadString8() & Salt)
        
        If LenB(newPass) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Debes especificar una contrasena nueva, intentalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
        Else
            storedPass = GetUserPassword(UserList(UserIndex).Name)
            
            If storedPass <> oldPass Then
                Call WriteConsoleMsg(UserIndex, "La contrasena actual proporcionada no es correcta. La contrasena no ha sido cambiada, intentalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call StorePasswordSalt(UserList(UserIndex).Name, newPass, Salt)
                Call WriteConsoleMsg(UserIndex, "La contrasena fue cambiada con exito.", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '10/07/2010: ZaMa - Now normal npcs don't answer if asked to gamble.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Amount  As Integer
        
        Dim TypeNpc As eNPCType
        
        Amount = Reader.ReadInt16()
        
        ' Dead
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            
            'Validate target NPC
        ElseIf .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            
            ' Validate Distance
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            
            ' Validate NpcType
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            
            Dim TargetNpcType As eNPCType
            
            TargetNpcType = Npclist(.flags.TargetNPC).NPCtype
            
            ' Normal npcs don't speak
            If TargetNpcType <> eNPCType.Comun And TargetNpcType <> eNPCType.DRAGON And TargetNpcType <> eNPCType.Pretoriano Then
                Call WriteChatOverHead(UserIndex, "No tengo ningUn interes en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
            End If
            
            ' Validate amount
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El minimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            ' Validate amount
        ElseIf Amount > 5000 Then
            Call WriteChatOverHead(UserIndex, "El maximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            ' Validate user gold
        ElseIf .Stats.Gld < Amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
        Else
            
            If RandomNumber(1, 100) <= 47 Then
                .Stats.Gld = .Stats.Gld + Amount
                Call WriteChatOverHead(UserIndex, "Felicidades! Has ganado " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.Gld = .Stats.Gld - Amount
                Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                
            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(UserIndex)
            
        End If
        
    End With
    
End Sub

''
' Handles the "InquiryVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiryVote(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim opt     As Byte
        
        opt = Reader.ReadInt8
        
        Call WriteConsoleMsg(UserIndex, ConsultaPopular.doVotar(UserIndex, opt), FontTypeNames.FONTTYPE_GUILD)
        
    End With
    
End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Amount  As Long
        
        Amount = Reader.ReadInt32()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        If Amount > 0 And Amount <= .Stats.Banco Then
            .Stats.Banco = .Stats.Banco - Amount
            .Stats.Gld = .Stats.Gld + Amount
            Call WriteChatOverHead(UserIndex, "Tenes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
        End If
        
        Call WriteUpdateGold(UserIndex)
        Call WriteUpdateBankGold(UserIndex)
        
    End With
    
End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 09/28/2010
    ' 09/28/2010 C4b3z0n - Ahora la respuesta de los NPCs sino perteneces a ninguna faccion solo la hacen el Rey o el Demonio
    ' 05/17/06 - Maraxus
    '***************************************************
    
    Dim TalkToKing  As Boolean
    
    Dim TalkToDemon As Boolean
    
    Dim NpcIndex    As Integer
    
    With UserList(UserIndex)
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        ' Chequea si habla con el rey o el demonio. Puede salir sin hacerlo, pero si lo hace le reponden los npcs
        NpcIndex = .flags.TargetNPC
        
        If NpcIndex <> 0 Then
            
            ' Es rey o domonio
            If Npclist(NpcIndex).NPCtype = eNPCType.Noble Then
                
                'Rey
                If Npclist(NpcIndex).flags.Faccion = 0 Then
                    TalkToKing = True
                    ' Demonio
                Else
                    TalkToDemon = True
                    
                End If
                
            End If
            
        End If
        
        'Quit the Royal Army
        If .Faccion.ArmadaReal = 1 Then
            
            ' Si le pidio al demonio salir de la armada, este le responde.
            If TalkToDemon Then
                Call WriteChatOverHead(UserIndex, "Sal de aqui bufon!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
                
            Else
                
                ' Si le pidio al rey salir de la armada, le responde.
                If TalkToKing Then
                    Call WriteChatOverHead(UserIndex, "Seras bienvenido a las fuerzas imperiales si deseas regresar.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
                    
                End If
                
                Call ExpulsarFaccionReal(UserIndex, False)
                
            End If
            
            'Quit the Chaos Legion
        ElseIf .Faccion.FuerzasCaos = 1 Then
            
            ' Si le pidio al rey salir del caos, le responde.
            If TalkToKing Then
                Call WriteChatOverHead(UserIndex, "Sal de aqui maldito criminal!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                
                ' Si le pidio al demonio salir del caos, este le responde.
                If TalkToDemon Then
                    Call WriteChatOverHead(UserIndex, "Ya volveras arrastrandote.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
                    
                End If
                
                Call ExpulsarFaccionCaos(UserIndex, False)
                
            End If
            
            ' No es faccionario
        Else
            
            ' Si le hablaba al rey o demonio, le repsonden ellos
            'Corregido, solo si son en efecto el rey o el demonio, no cualquier NPC (C4b3z0n)
            If (TalkToDemon And criminal(UserIndex)) Or (TalkToKing And Not criminal(UserIndex)) Then        'Si se pueden unir a la faccion (status), son invitados
            Call WriteChatOverHead(UserIndex, "No perteneces a nuestra faccion. Si deseas unirte, di /ENLISTAR", Npclist(NpcIndex).Char.CharIndex, vbWhite)
        ElseIf (TalkToDemon And Not criminal(UserIndex)) Then
            Call WriteChatOverHead(UserIndex, "Sal de aqui bufon!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
        ElseIf (TalkToKing And criminal(UserIndex)) Then
            Call WriteChatOverHead(UserIndex, "Sal de aqui maldito criminal!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
        Else
            Call WriteConsoleMsg(UserIndex, "No perteneces a ninguna faccion!", FontTypeNames.FONTTYPE_FIGHT)
            
        End If
        
    End If
    
End With

End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Amount  As Long
        
        Amount = Reader.ReadInt32()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
            
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        'Calculamos la diferencia con el maximo de oro permitido el cual es el valor de LONG
        Dim RemainingAmountToMaximumGold As Long
        RemainingAmountToMaximumGold = 2147483647 - .Stats.Gld
        
        If .Stats.Banco >= 2147483647 And RemainingAmountToMaximumGold <= Amount Then
            Call WriteChatOverHead(UserIndex, "No puedes depositar el oro por que tendrias mas del maximo permitido (2147483647)", Npclist(.flags.TargetNPC).Char.CharIndex, vbRed)
            
        ElseIf Amount > 0 And Amount <= .Stats.Gld Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.Gld = .Stats.Gld - Amount
            Call WriteChatOverHead(UserIndex, "Tenes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Call WriteUpdateGold(UserIndex)
            Call WriteUpdateBankGold(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No tenes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
        End If
        
    End With
    
End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 14/11/2010
    '14/11/2010: ZaMa - Now denounces can be desactivated.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Text    As String
        
        Dim Msg     As String
        
        Text = Reader.ReadString8()
        
        If .flags.Silenciado = 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Text)
            
            Msg = LCase$(.Name) & " DENUNCIA: " & Text
            
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(Msg, FontTypeNames.FONTTYPE_GUILDMSG), True)
            
            Call Denuncias.Push(Msg, False)
            
            Call WriteConsoleMsg(UserIndex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildFundate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundate(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        If HasFound(.Name) Then
            Call WriteConsoleMsg(UserIndex, "Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub
            
        End If
        
        Call WriteShowGuildAlign(UserIndex)
        
    End With
    
End Sub

''
' Handles the "GuildFundation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundation(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim clanType As eClanType
        
        Dim Error    As String
        
        clanType = Reader.ReadInt8
        
        If HasFound(.Name) Then
            Call WriteConsoleMsg(UserIndex, "Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogCheating("El usuario " & .Name & " ha intentado fundar un clan ya habiendo fundado otro desde la IP " & .IP)
            Exit Sub
            
        End If
        
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
                Call WriteConsoleMsg(UserIndex, "Alineacion invalida.", FontTypeNames.FONTTYPE_GUILD)
                Exit Sub
                
        End Select
        
        If modGuilds.PuedeFundarUnClan(UserIndex, .FundandoGuildAlineacion, Error) Then
            Call WriteShowGuildFundationForm(UserIndex)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)
            
        End If
        
    End With
    
End Sub

''
' Handles the "PartyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyKick(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/05/09
    'Last Modification by: Marco Vanotti (Marco)
    '- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If UserPuedeEjecutarComandos(UserIndex) Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tUser)
            Else
                
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                    
                End If
                
                Call WriteConsoleMsg(UserIndex, LCase(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "PartySetLeader" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartySetLeader(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/05/09
    'Last Modification by: Marco Vanotti (MarKoxX)
    '- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
    '***************************************************
    
    'On Error GoTo ErrHandler
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        Dim rank     As Integer
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.consejero
        
        UserName = Reader.ReadString8()
        
        If UserPuedeEjecutarComandos(UserIndex) Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                
                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(UserName) And rank) <= (.flags.Privilegios And rank) Then
                    Call mdParty.TransformarEnLider(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, LCase(UserList(tUser).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Else
                
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                    
                End If
                
                Call WriteConsoleMsg(UserIndex, LCase(UserName) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "PartyAcceptMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyAcceptMember(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/05/09
    'Last Modification by: Marco Vanotti (Marco)
    '- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName  As String
        
        Dim tUser     As Integer
        
        Dim rank      As Integer
        
        Dim bUserVivo As Boolean
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.consejero
        
        UserName = Reader.ReadString8()
        
        If UserList(UserIndex).flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_PARTY)
        Else
            bUserVivo = True
            
        End If
        
        If mdParty.UserPuedeEjecutarComandos(UserIndex) And bUserVivo Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                
                'Validate administrative ranks - don't allow users to spoof online GMs
                If (UserList(tUser).flags.Privilegios And rank) <= (.flags.Privilegios And rank) Then
                    Call mdParty.AprobarIngresoAParty(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes incorporar a tu party a personajes de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Else
                
                If InStr(UserName, "+") Then
                    UserName = Replace(UserName, "+", " ")
                    
                End If
                
                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(UserName) And rank) <= (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(UserIndex, LCase(UserName) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes incorporar a tu party a personajes de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GuildMemberList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Guild       As String
        
        Dim memberCount As Integer
        
        Dim i           As Long
        
        Dim UserName    As String
        
        Guild = Reader.ReadString8()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(Guild, "\") <> 0) Then
                Guild = Replace(Guild, "\", "")
                
            End If
            
            If (InStrB(Guild, "/") <> 0) Then
                Guild = Replace(Guild, "/", "")
                
            End If
            
            If Not FileExist(App.Path & "\guilds\" & Guild & "-members.mem") Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & Guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & Guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To memberCount
                    UserName = GetVar(App.Path & "\Guilds\" & Guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(UserIndex, UserName & "<" & Guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Message As String
        
        Message = Reader.ReadString8()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Mensaje a Gms:" & Message)
            
            If LenB(Message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(Message)
                
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & "> " & Message, FontTypeNames.FONTTYPE_GMMSG))
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName        'Show / Hide the name
            
            Call RefreshCharStatus(UserIndex)
            
        End If
        
    End With
    
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 28/05/2010
    '28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim i       As Long
        
        Dim list    As String
        
        Dim priv    As PlayerType
        
        priv = PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios
        
        ' Solo dioses pueden ver otros dioses online
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = priv Or PlayerType.Dios Or PlayerType.Admin
            
        End If
        
        For i = 1 To LastUser
            
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.ArmadaReal = 1 Then
                    If UserList(i).flags.Privilegios And priv Then
                        list = list & UserList(i).Name & ", "
                        
                    End If
                    
                End If
                
            End If
            
        Next i
        
    End With
    
    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Reales conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay reales conectados.", FontTypeNames.FONTTYPE_INFO)
        
    End If
    
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 28/05/2010
    '28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim i       As Long
        
        Dim list    As String
        
        Dim priv    As PlayerType
        
        priv = PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios
        
        ' Solo dioses pueden ver otros dioses online
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = priv Or PlayerType.Dios Or PlayerType.Admin
            
        End If
        
        For i = 1 To LastUser
            
            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.FuerzasCaos = 1 Then
                    If UserList(i).flags.Privilegios And priv Then
                        list = list & UserList(i).Name & ", "
                        
                    End If
                    
                End If
                
            End If
            
        Next i
        
    End With
    
    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Caos conectados.", FontTypeNames.FONTTYPE_INFO)
        
    End If
    
End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/10/07
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        UserName = Reader.ReadString8()
        
        Dim tIndex  As Integer
        
        Dim X       As Long
        
        Dim Y       As Long
        
        Dim i       As Long
        
        Dim Found   As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.consejero) Then
            
            'Si es dios o Admins no podemos salvo que nosotros tambien lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then        'existe el usuario destino
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                
                For i = 2 To 5        'esto for sirve ir cambiando la distancia destino
                    For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                        For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i
                            
                            If MapData(UserList(tIndex).Pos.Map, X, Y).UserIndex = 0 Then
                                If LegalPos(UserList(tIndex).Pos.Map, X, Y, True, True) Then
                                    Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, X, Y, True)
                                    Call LogGM(.Name, "/IRCERCA " & UserName & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
                                    Found = True
                                    Exit For
                                    
                                End If
                                
                            End If
                            
                        Next Y
                        
                        If Found Then Exit For        ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next X
                    
                    If Found Then Exit For        ' Feo, pero hay que abortar 3 fors sin usar GoTo
                Next i
                
                'No space found
                If Not Found Then
                    Call WriteConsoleMsg(UserIndex, "Todos los lugares estan ocupados.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End If
    
    'If we got here then packet is complete
    
End With

errHandler:

Dim Error           As Long

Error = Err.Number

On Error GoTo 0

If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim comment As String
        
        comment = Reader.ReadString8()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Comentario: " & comment)
            Call WriteConsoleMsg(UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call LogGM(.Name, "Hora.")
        
    End With
    
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & Time & " " & Date, FontTypeNames.FONTTYPE_INFO))
    
End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 18/11/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '18/11/2010: ZaMa - Obtengo los privs del charfile antes de mostrar la posicion de un usuario offline.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        Dim miPos    As String
        
        UserName = Reader.ReadString8()
        
        If Not .flags.Privilegios And PlayerType.User Then
            
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                
                If PersonajeExiste(UserName) Then
                    
                    Dim CharPrivs As PlayerType
                    
                    CharPrivs = GetCharPrivs(UserName)
                    
                    If (CharPrivs And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios)) <> 0 Or ((CharPrivs And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                        miPos = GetUserPos(UserName)
                        Call WriteConsoleMsg(UserIndex, "Ubicacion  " & UserName & " (Offline): " & miPos & ".", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                Else
                    
                    If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                End If
                
            Else
                
                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(UserIndex, "Ubicacion  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
        Call LogGM(.Name, "/Donde " & UserName)
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 30/07/06
    'Pablo (ToxicWaste): modificaciones generales para simplificar la visualizacion.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Map     As Integer
        
        Dim i, j    As Long
        
        Dim NPCcount1, NPCcount2 As Integer
        
        Dim NPCcant1() As Integer
        
        Dim NPCcant2() As Integer
        
        Dim List1()    As String
        
        Dim List2()    As String
        
        Map = Reader.ReadInt16()
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        If MapaValido(Map) Then
            
            For i = 1 To LastNPC
                
                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If Npclist(i).Pos.Map = Map Then
                    
                    'esta vivo
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else
                            
                            For j = 0 To NPCcount1 - 1
                                
                                If Left$(List1(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
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
                        
                    Else
                        
                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else
                            
                            For j = 0 To NPCcount2 - 1
                                
                                If Left$(List2(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
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
            
            Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            
            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles.", FontTypeNames.FONTTYPE_INFO)
            Else
                
                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j
                
            End If
            
            Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            
            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay mas NPCS.", FontTypeNames.FONTTYPE_INFO)
            Else
                
                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j
                
            End If
            
            Call LogGM(.Name, "Numero enemigos en mapa " & Map)
            
        End If
        
    End With
    
End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/09
    '26/03/06: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
    '***************************************************
    With UserList(UserIndex)
        
        Dim X       As Integer
        
        Dim Y       As Integer
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        X = .flags.TargetX
        Y = .flags.TargetY
        
        Call FindLegalPos(UserIndex, .flags.TargetMap, X, Y)
        Call WarpUserChar(UserIndex, .flags.TargetMap, X, Y, True)
        Call LogGM(.Name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.Map)
        
    End With
    
End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/08/2019
    '26/03/2009: ZaMa - Chequeo que no se teletransporte a un tile donde haya un char o npc.
    '11/08/2019: Jopi - No registramos en los logs si te teletransportas a vos mismo.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim Map      As Integer
        
        Dim X        As Integer
        
        Dim Y        As Integer
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        Map = Reader.ReadInt16()
        X = Reader.ReadInt8
        Y = Reader.ReadInt8
        
        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(Map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.consejero Then
                        tUser = NameIndex(UserName)
                        
                    End If
                    
                Else
                    tUser = UserIndex
                    
                End If
                
                If tUser <= 0 Then
                    If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                ElseIf Not ((UserList(tUser).flags.Privilegios And PlayerType.Dios) <> 0 Or (UserList(tUser).flags.Privilegios And PlayerType.Admin) <> 0) Or tUser = UserIndex Then
                    
                    If InMapBounds(Map, X, Y) Then
                        Call FindLegalPos(tUser, Map, X, Y)
                        Call WarpUserChar(tUser, Map, X, Y, True, True)
                        
                        ' Agrego esto para no llenar consola de mensajes al hacer SHIFT + CLICK DERECHO
                        If UserIndex <> tUser Then
                            Call WriteConsoleMsg(UserIndex, UserList(tUser).Name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Transporto a " & UserList(tUser).Name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)
                            
                        End If
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                
                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "Estimado usuario, ud. ha sido silenciado por los administradores. Sus denuncias seran ignoradas por el servidor de aqui en mas. Utilice /GM para contactar un administrador.")
                    Call LogGM(.Name, "/silenciar " & UserList(tUser).Name)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/DESsilenciar " & UserList(tUser).Name)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(UserIndex)
        
    End With
    
End Sub

''
' Handles the "RequestPartyForm" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyForm(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .PartyIndex > 0 Then
            Call WriteShowPartyForm(UserIndex)
            
        Else
            Call WriteConsoleMsg(UserIndex, "No perteneces a ningun grupo!", FontTypeNames.FONTTYPE_INFOBOLD)
            
        End If
        
    End With
    
End Sub

''
' Handles the "ItemUpgrade" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleItemUpgrade(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Torres Patricio
    'Last Modification: 12/09/09
    '
    '***************************************************
    With UserList(UserIndex)
        
        Dim ItemIndex As Integer
        
        ItemIndex = Reader.ReadInt16()
        
        If ItemIndex <= 0 Then Exit Sub
        If Not TieneObjetos(ItemIndex, 1, UserIndex) Then Exit Sub
        
        If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
        Call DoUpgrade(UserIndex, ItemIndex)
        
    End With
    
End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        UserName = Reader.ReadString8()
        
        If Not .flags.Privilegios And PlayerType.User Then Call Ayuda.Quitar(UserName)
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/2009
    '26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        Dim X        As Integer
        
        Dim Y        As Integer
        
        UserName = Reader.ReadString8()
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.consejero) Then
            
            'Si es dios o Admins no podemos salvo que nosotros tambien lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    X = UserList(tUser).Pos.X
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, X, Y)
                    
                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, X, Y, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                    Call LogGM(.Name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call DoAdminInvisible(UserIndex)
        Call LogGM(.Name, "/INVISIBLE")
        
    End With
    
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WriteShowGMPanelForm(UserIndex)
        
    End With
    
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/09/07
    'Last modified by: Lucas Tavolaro Ortiz (Tavo)
    'I haven`t found a solution to split, so i make an array of names
    '***************************************************
    Dim i           As Long
    
    Dim names()     As String
    
    Dim Count       As Long
    
    With UserList(UserIndex)
        
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
        
        If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)
        
    End With
    
End Sub

''
' Handles the "Working" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/10/2010
    '***************************************************
    Dim i           As Long
    
    Dim Users       As String
    
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            
            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                Users = Users & ", " & UserList(i).Name
                
            End If
            
        Next i
        
        If LenB(Users) <> 0 Then
            Users = Right$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios trabajando.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handles the "Hiding" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i           As Long
    
    Dim Users       As String
    
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            
            If (LenB(UserList(i).Name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                Users = Users & UserList(i).Name & ", "
                
            End If
            
        Next i
        
        If LenB(Users) <> 0 Then
            Users = Left$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios ocultandose.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    'Last Modification: 04/04/2020
    '4/4/2020: FrankoH298 - Ahora calcula bien el tiempo de carcel
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim Reason   As String
        
        Dim jailTime As Byte
        
        Dim Count    As Byte
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        Reason = Reader.ReadString8()
        jailTime = Reader.ReadInt8
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
            
        End If
        
        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser <= 0 Then
                    If (EsDios(UserName) Or EsAdmin(UserName)) Then
                        Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                Else
                    
                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > (60) Then
                        Call WriteConsoleMsg(UserIndex, "No puedes encarcelar por mas de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        
                        If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                            
                        End If
                        
                        If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                            
                        End If
                        
                        If PersonajeExiste(UserName) Then
                            Count = GetUserAmountOfPunishments(UserName)
                            Call SaveUserPunishment(UserName, Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & Time)
                            
                        End If
                        
                        Call Encarcelar(tUser, jailTime, .Name)
                        Call LogGM(.Name, " encarcelo a " & UserName)
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/22/08 (NicoNZ)
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim tNPC    As Integer
        
        Dim auxNPC  As npc
        
        #If ProteccionGM = 1 Then
            ' WyroX: A pedido de la gente, desactivo el comando /RMATA
            Call WriteConsoleMsg(UserIndex, "El comando /RMATA se encuentra desactivado.", FONTTYPE_SERVER)
            Exit Sub
        #End If
        
        'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
        If .flags.Privilegios And PlayerType.consejero Then
            If .Pos.Map = MAPA_PRETORIANO Then
                Call WriteConsoleMsg(UserIndex, "Los consejeros no pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
            End If
            
        End If
        
        tNPC = .flags.TargetNPC
        
        If tNPC > 0 Then
            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNPC).Name, FontTypeNames.FONTTYPE_INFO)
            
            auxNPC = Npclist(tNPC)
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
            
            .flags.TargetNPC = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Antes debes hacer click sobre el NPC.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/26/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim Reason   As String
        
        Dim Privs    As PlayerType
        
        Dim Count    As Byte
        
        UserName = Reader.ReadString8()
        Reason = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                Privs = UserDarPrivilegioLevel(UserName)
                
                If Not Privs And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "No puedes advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else
                    
                    If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                        
                    End If
                    
                    If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                        
                    End If
                    
                    If PersonajeExiste(UserName) Then
                        Count = GetUserAmountOfPunishments(UserName)
                        Call SaveUserPunishment(UserName, Count + 1, LCase$(.Name) & ": ADVERTENCIA por: " & LCase$(Reason) & " " & Date & " " & Time)
                        
                        Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(UserName) & ".", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, " advirtio a " & UserName)
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "EditChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 11/05/2019
    '02/03/2009: ZaMa - Cuando editas nivel, chequea si el pj puede permanecer en clan faccionario
    '11/06/2009: ZaMa - Todos los comandos se pueden usar aunque el pj este offline
    '18/09/2010: ZaMa - Ahora se puede editar la vida del propio pj (cualquier rm o dios).
    '11/05/2019: Jopi - No registramos en los logs si te editas a vos mismo.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName      As String
        
        Dim tUser         As Integer
        
        Dim opcion        As Byte
        
        Dim Arg1          As String
        
        Dim Arg2          As String
        
        Dim valido        As Boolean
        
        Dim LoopC         As Byte
        
        Dim CommandString As String
        
        Dim n             As Byte
        
        Dim UserCharPath  As String
        
        Dim Var           As Long
        
        UserName = Replace(Reader.ReadString8(), "+", " ")
        
        If UCase$(UserName) = "YO" Then
            tUser = UserIndex
        Else
            tUser = NameIndex(UserName)
            
        End If
        
        opcion = Reader.ReadInt8
        Arg1 = Reader.ReadString8()
        Arg2 = Reader.ReadString8()
        
        If .flags.Privilegios And PlayerType.RoleMaster Then
            
            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.consejero)
                
                Case PlayerType.consejero
                    ' Los RMs consejeros solo se pueden editar su head, body, level y vida
                    valido = tUser = UserIndex And (opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_Level Or opcion = eEditOptions.eo_Vida)
                    
                Case PlayerType.SemiDios
                    ' Los RMs solo se pueden editar su level o vida y el head y body de cualquiera
                    valido = ((opcion = eEditOptions.eo_Level Or opcion = eEditOptions.eo_Vida) And tUser = UserIndex) Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head
                    
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level o vida solo lo puede hacer sobre si mismo
                    valido = ((opcion = eEditOptions.eo_Level Or opcion = eEditOptions.eo_Vida) And tUser = UserIndex) Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_CiticensKilled Or opcion = eEditOptions.eo_CriminalsKilled Or opcion = eEditOptions.eo_Class Or opcion = eEditOptions.eo_Skills Or opcion = eEditOptions.eo_addGold
                    
            End Select
            
            'Si no es RM debe ser dios para poder usar este comando
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            
            If opcion = eEditOptions.eo_Vida Then
                '  Por ahora dejo para que los dioses no puedan editar la vida de otros
                valido = (tUser = UserIndex)
            Else
                valido = True
                
            End If
            
        ElseIf .flags.PrivEspecial Then
            valido = (opcion = eEditOptions.eo_CiticensKilled) Or (opcion = eEditOptions.eo_CriminalsKilled)
            
        End If
        
        'CHOTS | The user is not online and we are working with Database
        If Database_Enabled And tUser <= 0 Then
            valido = False
            Call WriteConsoleMsg(UserIndex, "El usuario esta offline.", FontTypeNames.FONTTYPE_INFO)
            
            '@TODO call a method to edit the user using the database
        End If
        
        If valido Then
            UserCharPath = CharPath & UserName & ".chr"
            
            If tUser <= 0 And Not FileExist(UserCharPath) Then
                Call WriteConsoleMsg(UserIndex, "Estas intentando editar un usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                Call LogGM(.Name, "Intento editar un usuario inexistente.")
            Else
                'For making the Log
                CommandString = "/MOD "
                
                Select Case opcion
                    
                    Case eEditOptions.eo_Gold
                        
                        If val(Arg1) <= MAX_ORO_EDIT Then
                            If tUser <= 0 Then        ' Esta offline
                            Call WriteVar(UserCharPath, "STATS", "GLD", val(Arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else        ' Online
                            UserList(tUser).Stats.Gld = val(Arg1)
                            Call WriteUpdateGold(tUser)
                            
                        End If
                        
                    Else
                        Call WriteConsoleMsg(UserIndex, "No esta permitido utilizar valores mayores a " & MAX_ORO_EDIT & ". Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                    ' Log it
                    CommandString = CommandString & "ORO "
                    
                Case eEditOptions.eo_Experience
                    
                    If val(Arg1) > 20000000 Then
                        Arg1 = 20000000
                        
                    End If
                    
                    If tUser <= 0 Then        ' Offline
                    Var = GetVar(UserCharPath, "STATS", "EXP")
                    Call WriteVar(UserCharPath, "STATS", "EXP", Var + val(Arg1))
                    Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else        ' Online
                    UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                    Call CheckUserLevel(tUser)
                    Call WriteUpdateExp(tUser)
                    
                End If
                
                ' Log it
                CommandString = CommandString & "EXP "
                
            Case eEditOptions.eo_Body
                
                If tUser <= 0 Then
                    Call WriteVar(UserCharPath, "INIT", "Body", Arg1)
                    Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                    
                End If
                
                ' Log it
                CommandString = CommandString & "BODY "
                
            Case eEditOptions.eo_Head
                
                If tUser <= 0 Then
                    Call WriteVar(UserCharPath, "INIT", "Head", Arg1)
                    Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call ChangeUserChar(tUser, UserList(tUser).Char.body, val(Arg1), UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                    
                End If
                
                ' Log it
                CommandString = CommandString & "HEAD "
                
            Case eEditOptions.eo_CriminalsKilled
                Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                
                If tUser <= 0 Then        ' Offline
                Call WriteVar(UserCharPath, "FACCIONES", "CrimMatados", Var)
                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
            Else        ' Online
                UserList(tUser).Faccion.CriminalesMatados = Var
                
            End If
            
            ' Log it
            CommandString = CommandString & "CRI "
            
        Case eEditOptions.eo_CiticensKilled
            Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
            
            If tUser <= 0 Then        ' Offline
            Call WriteVar(UserCharPath, "FACCIONES", "CiudMatados", Var)
            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
        Else        ' Online
            UserList(tUser).Faccion.CiudadanosMatados = Var
            
        End If
        
        ' Log it
        CommandString = CommandString & "CIU "
        
    Case eEditOptions.eo_Level
        
        If val(Arg1) > STAT_MAXELV Then
            Arg1 = CStr(STAT_MAXELV)
            Call WriteConsoleMsg(UserIndex, "No puedes tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)
            
        End If
        
        ' Chequeamos si puede permanecer en el clan
        If val(Arg1) >= 25 Then
            
            Dim GI  As Integer
            
            If tUser <= 0 Then
                GI = GetVar(UserCharPath, "GUILD", "GUILDINDEX")
            Else
                GI = UserList(tUser).GuildIndex
                
            End If
            
            If GI > 0 Then
                If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                    'We get here, so guild has factionary alignment, we have to expulse the user
                    Call modGuilds.m_EcharMiembroDeClan(-1, UserName)
                    
                    Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(UserName & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
                    
                    ' Si esta online le avisamos
                    If tUser > 0 Then Call WriteConsoleMsg(tUser, "Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearas! Por esta razon, hasta tanto no te enlistes en la faccion bajo la cual tu clan esta alineado, estaras excluido del mismo.", FontTypeNames.FONTTYPE_GUILD)
                    
                End If
                
            End If
            
        End If
        
        If tUser <= 0 Then        ' Offline
        Call WriteVar(UserCharPath, "STATS", "ELV", val(Arg1))
        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
    Else        ' Online
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
        Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
    Else
        
        If tUser <= 0 Then        ' Offline
        Call WriteVar(UserCharPath, "INIT", "Clase", LoopC)
        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
    Else        ' Online
        UserList(tUser).Clase = LoopC
        
    End If
    
End If

' Log it
CommandString = CommandString & "CLASE "

Case eEditOptions.eo_Skills
    
    For LoopC = 1 To NUMSKILLS
        
        If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
    Next LoopC
    
    If LoopC > NUMSKILLS Then
        Call WriteConsoleMsg(UserIndex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
    Else
        
        If tUser <= 0 Then        ' Offline
        Call WriteVar(UserCharPath, "Skills", "SK" & LoopC, Arg2)
        Call WriteVar(UserCharPath, "Skills", "EXPSK" & LoopC, 0)
        
        If Arg2 < MAXSKILLPOINTS Then
            Call WriteVar(UserCharPath, "Skills", "ELUSK" & LoopC, ELU_SKILL_INICIAL * 1.05 ^ Arg2)
        Else
            Call WriteVar(UserCharPath, "Skills", "ELUSK" & LoopC, 0)
            
        End If
        
        Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
    Else        ' Online
        UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)
        Call CheckEluSkill(tUser, LoopC, True)
        
    End If
    
End If

' Log it
CommandString = CommandString & "SKILLS "

Case eEditOptions.eo_SkillPointsLeft
    
    If tUser <= 0 Then        ' Offline
    Call WriteVar(UserCharPath, "STATS", "SkillPtsLibres", Arg1)
    Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
Else        ' Online
    UserList(tUser).Stats.SkillPts = val(Arg1)
    
End If

' Log it
CommandString = CommandString & "SKILLSLIBRES "

Case eEditOptions.eo_Nobleza
    Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
    
    If tUser <= 0 Then        ' Offline
    Call WriteVar(UserCharPath, "REP", "Nobles", Var)
    Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
Else        ' Online
    UserList(tUser).Reputacion.NobleRep = Var
    
End If

' Log it
CommandString = CommandString & "NOB "

Case eEditOptions.eo_Asesino
    Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
    
    If tUser <= 0 Then        ' Offline
    Call WriteVar(UserCharPath, "REP", "Asesino", Var)
    Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
Else        ' Online
    UserList(tUser).Reputacion.AsesinoRep = Var
    
End If

' Log it
CommandString = CommandString & "ASE "

Case eEditOptions.eo_Sex
    
    Dim Sex         As Byte
    
    Sex = IIf(UCase(Arg1) = "MUJER", eGenero.Mujer, 0)        ' Mujer
    Sex = IIf(UCase(Arg1) = "HOMBRE", eGenero.Hombre, Sex)        ' Hombre
    
    If Sex <> 0 Then        ' Es Hombre o mujer
    If tUser <= 0 Then        ' OffLine
    Call WriteVar(UserCharPath, "INIT", "Genero", Sex)
    Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
Else        ' Online
    UserList(tUser).Genero = Sex
    
End If

Else
    Call WriteConsoleMsg(UserIndex, "Genero desconocido. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
    
End If

' Log it
CommandString = CommandString & "SEX "

Case eEditOptions.eo_Raza
    
    Dim raza        As Byte
    
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
        Call WriteConsoleMsg(UserIndex, "Raza desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
    Else
        
        If tUser <= 0 Then
            Call WriteVar(UserCharPath, "INIT", "Raza", raza)
            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
        Else
            UserList(tUser).raza = raza
            
        End If
        
    End If
    
    ' Log it
    CommandString = CommandString & "RAZA "
    
Case eEditOptions.eo_addGold
    
    Dim bankGold    As Long
    
    If Abs(Arg1) > MAX_ORO_EDIT Then
        Call WriteConsoleMsg(UserIndex, "No esta permitido utilizar valores mayores a " & MAX_ORO_EDIT & ".", FontTypeNames.FONTTYPE_INFO)
    Else
        
        If tUser <= 0 Then
            bankGold = GetVar(UserCharPath, "STATS", "BANCO")
            Call WriteVar(UserCharPath, "STATS", "BANCO", IIf(bankGold + val(Arg1) <= 0, 0, bankGold + val(Arg1)))
            Call WriteConsoleMsg(UserIndex, "Se le ha agregado " & Arg1 & " monedas de oro a " & UserName & ".", FONTTYPE_TALK)
        Else
            UserList(tUser).Stats.Banco = IIf(UserList(tUser).Stats.Banco + val(Arg1) <= 0, 0, UserList(tUser).Stats.Banco + val(Arg1))
            Call WriteConsoleMsg(tUser, STANDARD_BOUNTY_HUNTER_MESSAGE, FONTTYPE_TALK)
            
        End If
        
    End If
    
    ' Log it
    CommandString = CommandString & "AGREGAR "
    
Case eEditOptions.eo_Vida
    
    If val(Arg1) > MAX_VIDA_EDIT Then
        Arg1 = CStr(MAX_VIDA_EDIT)
        Call WriteConsoleMsg(UserIndex, "No puedes tener vida superior a " & MAX_VIDA_EDIT & ".", FONTTYPE_INFO)
        
    End If
    
    ' No valido si esta offline, porque solo se puede editar a si mismo
    UserList(tUser).Stats.MaxHp = val(Arg1)
    UserList(tUser).Stats.MinHp = val(Arg1)
    
    Call WriteUpdateUserStats(tUser)
    
    ' Log it
    CommandString = CommandString & "VIDA "
    
Case eEditOptions.eo_Poss
    
    Dim Map         As Integer
    
    Dim X           As Integer
    
    Dim Y           As Integer
    
    Map = val(ReadField(1, Arg1, 45))
    X = val(ReadField(2, Arg1, 45))
    Y = val(ReadField(3, Arg1, 45))
    
    If InMapBounds(Map, X, Y) Then
        
        If tUser <= 0 Then
            Call WriteVar(UserCharPath, "INIT", "POSITION", Map & "-" & X & "-" & Y)
            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WarpUserChar(tUser, Map, X, Y, True, True)
            Call WriteConsoleMsg(UserIndex, "Usuario teletransportado: " & UserName, FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    Else
        Call WriteConsoleMsg(UserIndex, "Posicion invalida", FONTTYPE_INFO)
        
    End If
    
    ' Log it
    CommandString = CommandString & "POSS "
    
Case Else
    Call WriteConsoleMsg(UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
    CommandString = CommandString & "UNKOWN "
    
End Select

CommandString = CommandString & Arg1 & " " & Arg2

If UserIndex <> tUser Then
    Call LogGM(.Name, CommandString & " " & UserName)
End If

End If

End If

'If we got here then packet is complete
End With

errHandler:

Dim Error           As Long

Error = Err.Number

On Error GoTo 0

If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid).. alto bug zapallo..
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim TargetName  As String
        
        Dim targetIndex As Integer
        
        TargetName = Replace$(Reader.ReadString8(), "+", " ")
        targetIndex = NameIndex(TargetName)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            
            'is the player offline
            If targetIndex <= 0 Then
                
                'don't allow to retrieve administrator's info
                If Not (EsDios(TargetName) Or EsAdmin(TargetName)) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, buscando...", FontTypeNames.FONTTYPE_INFO)
                    
                    If Not Database_Enabled Then
                        Call SendUserStatsTxtCharfile(UserIndex, TargetName)
                    Else
                        Call SendUserStatsTxtDatabase(UserIndex, TargetName)
                        
                    End If
                    
                End If
                
            Else
                
                'don't allow to retrieve administrator's info
                If UserList(targetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(UserIndex, targetIndex)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName         As String
        
        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean
        
        Dim OtherUserIsAdmin As Boolean
        
        UserName = Reader.ReadString8()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And ((.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin) Then
            Call LogGM(.Name, "/STAT " & UserName)
            
            tUser = NameIndex(UserName)
            
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Buscando... ", FontTypeNames.FONTTYPE_INFO)
                    
                    If Not Database_Enabled Then
                        Call SendUserMiniStatsTxtFromCharfile(UserIndex, UserName)
                    Else
                        Call SendUserMiniStatsTxtFromDatabase(UserIndex, UserName)
                        
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver los stats de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Else
                
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserMiniStatsTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver los stats de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName         As String
        
        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean
        
        Dim OtherUserIsAdmin As Boolean
        
        UserName = Reader.ReadString8()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) Or UserIsAdmin Then
            
            Call LogGM(.Name, "/BAL " & UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Buscando... ", FontTypeNames.FONTTYPE_TALK)
                    
                    If Not Database_Enabled Then
                        Call SendUserOROTxtFromCharfile(UserIndex, UserName)
                    Else
                        Call SendUserOROTxtFromDatabase(UserIndex, UserName)
                        
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el oro de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Else
                
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco.", FontTypeNames.FONTTYPE_TALK)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el oro de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName         As String
        
        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean
        
        Dim OtherUserIsAdmin As Boolean
        
        UserName = Reader.ReadString8()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/INV " & UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Buscando...", FontTypeNames.FONTTYPE_TALK)
                    
                    If Not Database_Enabled Then
                        Call SendUserInvTxtFromCharfile(UserIndex, UserName)
                    Else
                        Call SendUserInvTxtFromDatabase(UserIndex, UserName)
                        
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Else
                
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserInvTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName         As String
        
        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean
        
        Dim OtherUserIsAdmin As Boolean
        
        UserName = Reader.ReadString8()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin Then
            Call LogGM(.Name, "/BOV " & UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            tUser = NameIndex(UserName)
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Buscando... ", FontTypeNames.FONTTYPE_TALK)
                    
                    If Not Database_Enabled Then
                        Call SendUserBovedaTxtFromCharfile(UserIndex, UserName)
                    Else
                        Call SendUserBovedaTxtFromDatabase(UserIndex, UserName)
                        
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver la boveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Else
                
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserBovedaTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver la boveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        Dim LoopC    As Long
        
        Dim Message  As String
        
        UserName = Reader.ReadString8()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/STATS " & UserName)
            
            If tUser <= 0 Then
                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", "")
                    
                End If
                
                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", "")
                    
                End If
                
                For LoopC = 1 To NUMSKILLS
                    Message = Message & GetUserSkills(UserName)
                Next LoopC
                
                Call WriteConsoleMsg(UserIndex, Message & "CHAR> Libres: " & GetUserFreeSkills(UserName), FontTypeNames.FONTTYPE_INFO)
                
            Else
                Call SendUserSkillsTxt(UserIndex, tUser)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 11/03/2010
    '11/03/2010: ZaMa - Al revivir con el comando, si esta navegando le da cuerpo e barca.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        Dim LoopC    As Byte
        
        UserName = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = UserIndex
                
            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
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
                            Call WriteMultiMessage(tUser, eMessages.CancelHome)
                            
                        End If
                        
                        Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                    .Stats.MinHp = .Stats.MaxHp
                    
                    If .flags.Traveling = 1 Then
                        .Counters.goHome = 0
                        .flags.Traveling = 0
                        Call WriteMultiMessage(tUser, eMessages.CancelHome)
                        
                    End If
                    
                End With
                
                Call WriteUpdateHP(tUser)
                
                Call LogGM(.Name, "Resucito a " & UserName)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 12/28/06
    '
    '***************************************************
    Dim i           As Long
    
    Dim list        As String
    
    Dim priv        As PlayerType
    
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero) Then Exit Sub
        
        priv = PlayerType.consejero Or PlayerType.SemiDios
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        
        For i = 1 To LastUser
            
            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then list = list & UserList(i).Name & ", "
                
            End If
            
        Next i
        
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 23/03/2009
    '23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
    '***************************************************
    With UserList(UserIndex)
        
        Dim Map     As Integer
        
        Map = Reader.ReadInt16
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero) Then Exit Sub
        
        Dim LoopC   As Long
        
        Dim list    As String
        
        Dim priv    As PlayerType
        
        priv = PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)
        
        For LoopC = 1 To LastUser
            
            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.Map = Map Then
                If UserList(LoopC).flags.Privilegios And priv Then list = list & UserList(LoopC).Name & ", "
                
            End If
            
        Next LoopC
        
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
        Call LogGM(.Name, "/ONLINEMAP " & Map)
        
    End With
    
End Sub

''
' Handles the "Forgive" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForgive(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If EsNewbie(tUser) Then
                    Call VolverCiudadano(tUser)
                Else
                    Call LogGM(.Name, "Intento perdonar un personaje de nivel avanzado.")
                    
                    If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                        Call WriteConsoleMsg(UserIndex, "Solo se permite perdonar newbies.", FontTypeNames.FONTTYPE_INFO)
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        Dim rank     As Integer
        
        Dim IsAdmin  As Boolean
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.consejero
        
        UserName = Reader.ReadString8()
        IsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) Or IsAdmin Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Else
                
                If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " echo a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseSocket(tUser)
                    Call LogGM(.Name, "Echo a " & UserName)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "Estas loco?? Como vas a pinatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserName & ".", FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.Name, " ejecuto a " & UserName)
                    
                End If
                
            Else
                
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Then
                    Call WriteConsoleMsg(UserIndex, "No esta online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Estas loco?? Como vas a pinatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim Reason   As String
        
        UserName = Reader.ReadString8()
        Reason = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(UserIndex, UserName, Reason)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName  As String
        
        Dim cantPenas As Byte
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
                
            End If
            
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
                
            End If
            
            If Not PersonajeExiste(UserName) Then
                Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
            Else
                
                If BANCheck(UserName) Then
                    Call UnBan(UserName)
                    
                    'penas
                    cantPenas = GetUserAmountOfPunishments(UserName)
                    Call SaveUserPunishment(UserName, cantPenas + 1, LCase$(.Name) & ": UNBAN. " & Date & " " & Time)
                    
                    Call LogGM(.Name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(UserIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & " no esta baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero) Then Exit Sub
        
        #If ProteccionGM = 1 Then
            ' WyroX: A pedido de la gente, desactivo el comando /SEGUIR
            Call WriteConsoleMsg(UserIndex, "El comando /SEGUIR se encuentra desactivado.", FONTTYPE_SERVER)
            Exit Sub
        #End If
        
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

Private Sub HandleSummonChar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 26/03/2009
    '26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        Dim X        As Integer
        
        Dim Y        As Integer
        
        UserName = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                If EsDios(UserName) Or EsAdmin(UserName) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "El jugador no esta online.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Else
                
                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (UserList(tUser).flags.Privilegios And (PlayerType.consejero Or PlayerType.User)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                    X = .Pos.X
                    Y = .Pos.Y + 1
                    Call FindLegalPos(tUser, .Pos.Map, X, Y)
                    Call WarpUserChar(tUser, .Pos.Map, X, Y, True, True)
                    Call LogGM(.Name, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero) Then Exit Sub
        
        Call EnviarSpawnList(UserIndex)
        
    End With
    
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim npc     As Integer
        
        npc = Reader.ReadInt16()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)
            
            Call LogGM(.Name, "Sumoneo " & Declaraciones.SpawnList(npc).NpcName)
            
        End If
        
    End With
    
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.Name, "/RESETINV " & Npclist(.flags.TargetNPC).Name)
        
    End With
    
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 28/05/2010
    '28/05/2010: ZaMa - Ahora no dice el nombre del gm que lo dice.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Message As String
        
        Message = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & Message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & "> " & Message, FontTypeNames.FONTTYPE_GUILD))
                
                ''''''''''''''''SOLO PARA EL TESTEO'''''''
                ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
                'frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & UserList(UserIndex).name & " > " & message
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "MapMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMapMessage(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Message As String
        
        Message = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Message) <> 0 Then
                
                Dim Mapa As Integer
                Mapa = .Pos.Map
                
                Call LogGM(.Name, "Mensaje a mapa " & Mapa & ":" & Message)
                Call SendData(SendTarget.toMap, Mapa, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_TALK))
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    'Pablo (ToxicWaste): Agrego para que el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        Dim priv     As PlayerType
        
        Dim IsAdmin  As Boolean
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.Name, "NICK2IP Solicito la IP de " & UserName)
            
            IsAdmin = (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0
            
            If IsAdmin Then
                priv = PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                priv = PlayerType.User
                
            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).IP, FontTypeNames.FONTTYPE_INFO)
                    
                    Dim IP    As String
                    
                    Dim lista As String
                    
                    Dim LoopC As Long
                    
                    IP = UserList(tUser).IP
                    
                    For LoopC = 1 To LastUser
                        
                        If UserList(LoopC).IP = IP Then
                            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).Name & ", "
                                    
                                End If
                                
                            End If
                            
                        End If
                        
                    Next LoopC
                    
                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Else
                
                If Not (EsDios(UserName) Or EsAdmin(UserName)) Or IsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "No hay ningUn personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim IP      As String
        
        Dim LoopC   As Long
        
        Dim lista   As String
        
        Dim priv    As PlayerType
        
        IP = Reader.ReadInt8 & "."
        IP = IP & Reader.ReadInt8 & "."
        IP = IP & Reader.ReadInt8 & "."
        IP = IP & Reader.ReadInt8
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & IP)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User
            
        End If
        
        For LoopC = 1 To LastUser
            
            If UserList(LoopC).IP = IP Then
                If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).Name & ", "
                        
                    End If
                    
                End If
                
            End If
            
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim GuildName As String
        
        Dim tGuild    As Integer
        
        GuildName = Reader.ReadString8()
        
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")
            
        End If
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = GuildIndex(GuildName)
            
            If tGuild > 0 Then
                Call WriteConsoleMsg(UserIndex, "Clan " & UCase(GuildName) & ": " & modGuilds.m_ListaDeMiembrosOnline(UserIndex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 22/03/2010
    '15/11/2009: ZaMa - Ahora se crea un teleport con un radio especificado.
    '22/03/2010: ZaMa - Harcodeo los teleps y radios en el dat, para evitar mapas bugueados.
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Mapa    As Integer
        
        Dim X       As Byte
        
        Dim Y       As Byte
        
        Dim Radio   As Byte
        
        Mapa = Reader.ReadInt16()
        X = Reader.ReadInt8
        Y = Reader.ReadInt8
        Radio = Reader.ReadInt8
        
        Radio = MinimoInt(Radio, 6)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.Name, "/CT " & Mapa & "," & X & "," & Y & "," & Radio)
        
        If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then Exit Sub
        
        If MapData(Mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapData(Mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        Dim ET      As obj
        
        ET.Amount = 1
        ' Es el numero en el dat. El indice es el comienzo + el radio, todo harcodeado :(.
        ET.ObjIndex = TELEP_OBJ_INDEX + Radio
        
        With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
            .TileExit.Map = Mapa
            .TileExit.X = X
            .TileExit.Y = Y
            
        End With
        
        Call MakeObj(ET, .Pos.Map, .Pos.X, .Pos.Y - 1)
        
    End With
    
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        Dim Mapa    As Integer
        
        Dim X       As Byte
        
        Dim Y       As Byte
        
        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        With MapData(Mapa, X, Y)
            
            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
                
                Call LogGM(UserList(UserIndex).Name, "/DT: " & Mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.Amount, Mapa, X, Y)
                
                If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(1, .TileExit.Map, .TileExit.X, .TileExit.Y)
                    
                End If
                
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
                
            End If
            
        End With
        
    End With
    
End Sub

''
' Handles the "ExitDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExitDestroy(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Cucsifae
    'Last Modification: 30/9/18
    '
    '***************************************************
    With UserList(UserIndex)
        
        Dim Mapa    As Integer
        
        Dim X       As Byte
        
        Dim Y       As Byte
        
        '/de
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        With MapData(Mapa, X, Y)
            
            If .TileExit.Map = 0 Then Exit Sub
            
            'Si hay un Teleport hay que usar /DT
            If .ObjInfo.ObjIndex > 0 Then
                If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then Exit Sub
                
            End If
            
            Call LogGM(UserList(UserIndex).Name, "/DE: " & Mapa & "," & X & "," & Y)
            
            .TileExit.Map = 0
            .TileExit.X = 0
            .TileExit.Y = 0
            
        End With
        
    End With
    
End Sub

''
' Handles the "RainToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRainToggle(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero) Then Exit Sub
        
        Call LogGM(.Name, "/LLUVIA")
        Lloviendo = Not Lloviendo
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        
    End With
    
End Sub

''
' Handles the "EnableDenounces" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnableDenounces(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    'Enables/Disables
    '***************************************************
    
    With UserList(UserIndex)
        
        If Not EsGm(UserIndex) Then Exit Sub
        
        Dim Activado As Boolean
        
        Dim Msg      As String
        
        Activado = Not .flags.SendDenounces
        .flags.SendDenounces = Activado
        
        Msg = "Denuncias por consola " & IIf(Activado, "ativadas", "desactivadas") & "."
        
        Call LogGM(.Name, Msg)
        
        Call WriteConsoleMsg(UserIndex, Msg, FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handles the "ShowDenouncesList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowDenouncesList(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowDenounces(UserIndex)
        
    End With
    
End Sub

''
' Handles the "SetCharDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim tUser   As Integer
        
        Dim Desc    As String
        
        Desc = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser
            
            If tUser > 0 Then
                UserList(tUser).DescRM = Desc
            Else
                Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes.", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ForceMP3ToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMP3ToMap(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Recoaro(Recox)
    'Last Modification: 07/01/20
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Mp3Id   As Byte
        
        Dim Mapa    As Integer
        
        Mp3Id = Reader.ReadInt8
        Mapa = Reader.ReadInt16
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, 50, 50) Then
                Mapa = .Pos.Map
                
            End If
            
            If Mp3Id = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMp3(MapInfo(.Pos.Map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMp3(Mp3Id))
                
            End If
            
        End If
        
    End With
    
End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim midiID  As Byte
        
        Dim Mapa    As Integer
        
        midiID = Reader.ReadInt8
        Mapa = Reader.ReadInt16
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, 50, 50) Then
                Mapa = .Pos.Map
                
            End If
            
            If midiID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(MapInfo(.Pos.Map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(midiID))
                
            End If
            
        End If
        
    End With
    
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim waveID  As Byte
        
        Dim Mapa    As Integer
        
        Dim X       As Byte
        
        Dim Y       As Byte
        
        waveID = Reader.ReadInt8
        Mapa = Reader.ReadInt16()
        X = Reader.ReadInt8
        Y = Reader.ReadInt8
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, X, Y) Then
                Mapa = .Pos.Map
                X = .Pos.X
                Y = .Pos.Y
                
            End If
            
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, X, Y))
            
        End If
        
    End With
    
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Message As String
        
        Message = Reader.ReadString8()
        
        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("EJERCITO REAL> " & Message, FontTypeNames.FONTTYPE_TALK))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Message As String
        
        Message = Reader.ReadString8()
        
        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & Message, FontTypeNames.FONTTYPE_TALK))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "CitizenMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCitizenMessage(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Message As String
        
        Message = Reader.ReadString8()
        
        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & Message, FontTypeNames.FONTTYPE_TALK))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "CriminalMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCriminalMessage(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Message As String
        
        Message = Reader.ReadString8()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & Message, FontTypeNames.FONTTYPE_TALK))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Message As String
        
        Message = Reader.ReadString8()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            
            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(Message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X       As Long
        
        Dim Y       As Long
        
        Dim bIsExit As Boolean
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                        
                        If ItemNoEsDeMapa(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)
                            
                        End If
                        
                    End If
                    
                End If
                
            Next X
        Next Y
        
        Call LogGM(UserList(UserIndex).Name, "/MASSDEST")
        
    End With
    
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        Dim LoopC    As Byte
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                    
                End With
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        Dim LoopC    As Byte
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)
                    
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil
                    
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                    
                End With
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim tObj    As Integer
        
        Dim lista   As String
        
        Dim X       As Long
        
        Dim Y       As Long
        
        For X = 5 To 95
            For Y = 5 To 95
                tObj = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                If tObj > 0 Then
                    If ObjData(tObj).OBJType <> eOBJType.otArboles Then
                        Call WriteConsoleMsg(UserIndex, "(" & X & "," & Y & ") " & ObjData(tObj).Name, FontTypeNames.FONTTYPE_INFO)
                        
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

Private Sub HandleMakeDumb(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SecurityIp.DumpTables
        
    End With
    
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                If PersonajeExiste(UserName) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, echando de los consejos.", FontTypeNames.FONTTYPE_INFO)
                    Call KickUserCouncils(UserName)
                Else
                    Call WriteConsoleMsg(UserIndex, "No se encuentra el charfile " & CharPath & UserName, FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Else
                
                With UserList(tUser)
                    
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                        
                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                        
                    End If
                    
                End With
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim tTrigger As Byte
        
        Dim tLog     As String
        
        tTrigger = Reader.ReadInt8
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If tTrigger >= 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            
            Call LogGM(.Name, tLog)
            Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 04/13/07
    '
    '***************************************************
    Dim tTrigger    As Byte
    
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger
        
        Call LogGM(.Name, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(UserIndex, "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim lista   As String
        
        Dim LoopC   As Long
        
        Call LogGM(.Name, "/BANIPLIST")
        
        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call BanIpGuardar
        Call BanIpCargar
        
    End With
    
End Sub

''
' Handles the "GuildBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim GuildName   As String
        
        Dim cantMembers As Integer
        
        Dim LoopC       As Long
        
        Dim member      As String
        
        Dim Count       As Byte
        
        Dim tIndex      As Integer
        
        Dim tFile       As String
        
        GuildName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " baneo al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_GUILD))
                
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
                    
                    Call SaveBan(member, "BAN AL CLAN: " & GuildName, LCase$(.Name))
                Next LoopC
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/02/09
    'Agregado un CopyBuffer porque se producia un bucle
    'inifito al intentar banear una ip ya baneada. (NicoNZ)
    '07/02/09 Pato - Ahora no es posible saber si un gm esta o no online.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim bannedIP As String
        
        Dim tUser    As Integer
        
        Dim Reason   As String
        
        Dim i        As Long
        
        ' Is it by ip
        If Reader.ReadBool() Then
            bannedIP = Reader.ReadInt8 & "."
            bannedIP = bannedIP & Reader.ReadInt8 & "."
            bannedIP = bannedIP & Reader.ReadInt8 & "."
            bannedIP = bannedIP & Reader.ReadInt8
        Else
            tUser = NameIndex(Reader.ReadString8())
            
            If tUser > 0 Then bannedIP = UserList(tUser).IP
            
        End If
        
        Reason = Reader.ReadString8()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIP) > 0 Then
                Call LogGM(.Name, "/BanIP " & bannedIP & " por " & Reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIP)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " baneo la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))
                    
                    'Find every player with that ip and ban him!
                    For i = 1 To LastUser
                        
                        If UserList(i).ConnIDValida Then
                            If UserList(i).IP = bannedIP Then
                                Call BanCharacter(UserIndex, UserList(i).Name, "IP POR " & Reason)
                                
                            End If
                            
                        End If
                        
                    Next i
                    
                End If
                
            ElseIf tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no esta online.", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim bannedIP As String
        
        bannedIP = Reader.ReadInt8 & "."
        bannedIP = bannedIP & Reader.ReadInt8 & "."
        bannedIP = bannedIP & Reader.ReadInt8 & "."
        bannedIP = bannedIP & Reader.ReadInt8
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 11/02/2011
    'maTih.- : Ahora se puede elegir, la cantidad a crear.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        ' Recibo el ID del paquete
        
        Dim tObj    As Integer: tObj = Reader.ReadInt16()
        Dim Cuantos As Integer: Cuantos = Reader.ReadInt16()
        
        ' Es Game-Master
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        ' WyroX: A pedido de la gente, desactivo el comando /CI
        #If ProteccionGM = 1 Then
            Call WriteConsoleMsg(UserIndex, "El comando /CI se encuentra desactivado.", FONTTYPE_SERVER)
            Exit Sub
        #End If
        
        ' Si hace mas de 10000, lo sacamos cagando.
        If Cuantos > 10000 Then Call WriteConsoleMsg(UserIndex, "Estas tratando de crear demasiado, como mucho podes crear 10.000 unidades.", FontTypeNames.FONTTYPE_TALK): Exit Sub
        
        ' El indice proporcionado supera la cantidad minima o total de items existentes en el juego
        If tObj < 1 Or tObj > NumObjDatas Then Exit Sub
        
        ' El nombre del objeto es nulo
        If LenB(ObjData(tObj).Name) = 0 Then Exit Sub
        
        Dim Objeto  As obj
        
        With Objeto
            .Amount = Cuantos
            .ObjIndex = tObj
        End With
        
        ' Chequeo si el objeto es AGARRABLE(para las puertas, arboles y demas objs. que no deberian estar en el inventario)
        '   0 = SI
        '   1 = NO
        If ObjData(tObj).Agarrable = 0 Then
            ' Trato de meterlo en el inventario.
            If MeterItemEnInventario(UserIndex, Objeto) Then
                Call WriteConsoleMsg(UserIndex, "Has creado " & Objeto.Amount & " unidades de " & ObjData(tObj).Name & ".", FontTypeNames.FONTTYPE_INFO)
            Else
                ' Si no hay espacio, lo tiro al piso.
                Call TirarItemAlPiso(.Pos, Objeto)
                Call WriteConsoleMsg(UserIndex, "No tenes espacio en tu inventario para crear el item.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
            End If
        Else
            ' Crear el item NO AGARRARBLE y tirarlo al piso.
            Call TirarItemAlPiso(.Pos, Objeto)
            Call WriteConsoleMsg(UserIndex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        ' Lo registro en los logs.
        Call LogGM(.Name, "/CI: " & tObj & " - [Nombre del Objeto: " & ObjData(tObj).Name & "] - [Cantidad : " & Cuantos & "]")
        
    End With
    
errHandler:
    If Err.Number <> 0 Then
        Call LogError("Error en HandleCreateItem " & Err.Number & " " & Err.description)
    End If
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim Mapa    As Integer
        
        Dim X       As Byte
        
        Dim Y       As Byte
        
        Mapa = .Pos.Map
        X = .Pos.X
        Y = .Pos.Y
        
        Dim ObjIndex As Integer
        
        ObjIndex = MapData(Mapa, X, Y).ObjInfo.ObjIndex
        
        If ObjIndex = 0 Then Exit Sub
        
        Call LogGM(.Name, "/DEST " & ObjIndex & " en mapa " & Mapa & " (" & X & "," & Y & "). Cantidad: " & MapData(Mapa, X, Y).ObjInfo.Amount)
        
        If ObjData(ObjIndex).OBJType = eOBJType.otTeleport And MapData(Mapa, X, Y).TileExit.Map > 0 Then
            
            Call WriteConsoleMsg(UserIndex, "No puede destruir teleports asi. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        Call EraseObj(10000, Mapa, X, Y)
        
    End With
    
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or .flags.PrivEspecial Then
            
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
                
            End If
            
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
                
            End If
            
            tUser = NameIndex(UserName)
            
            Call LogGM(.Name, "ECHO DEL CAOS A: " & UserName)
            
            If tUser > 0 Then
                Call ExpulsarFaccionCaos(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                
                If PersonajeExiste(UserName) Then
                    Call KickUserChaosLegion(UserName, .Name)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & " inexistente.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or .flags.PrivEspecial Then
            
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
                
            End If
            
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
                
            End If
            
            tUser = NameIndex(UserName)
            
            Call LogGM(.Name, "ECHO DE LA REAL A: " & UserName)
            
            If tUser > 0 Then
                Call ExpulsarFaccionReal(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                
                If PersonajeExiste(UserName) Then
                    Call KickUserRoyalArmy(UserName, .Name)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & " inexistente.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ForceMP3All" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMP3All(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Recoaro(Recox)
    'Last Modification: 07/01/20
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim Mp3Id   As Byte
        
        Mp3Id = Reader.ReadInt8
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast musica MP3: " & Mp3Id, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMp3(Mp3Id))
        
    End With
    
End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim midiID  As Byte
        
        midiID = Reader.ReadInt8
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast musica MIDI: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))
        
    End With
    
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim waveID  As Byte
        
        waveID = Reader.ReadInt8
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
        
    End With
    
End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 1/05/07
    'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName   As String
        
        Dim punishment As Byte
        
        Dim NewText    As String
        
        UserName = Reader.ReadString8()
        punishment = Reader.ReadInt8
        NewText = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else
                
                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", "")
                    
                End If
                
                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", "")
                    
                End If
                
                If PersonajeExiste(UserName) Then
                    Call LogGM(.Name, " borro la pena: " & punishment & " de " & UserName & " y la cambio por: " & NewText)
                    
                    Call AlterUserPunishment(UserName, punishment, LCase$(.Name) & ": <" & NewText & "> " & Date & " " & Time)
                    Call WriteConsoleMsg(UserIndex, "Pena modificada.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.Name, "/BLOQ")
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0
            
        End If
        
        Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked)
        
    End With
    
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        #If ProteccionGM = 1 Then
            ' WyroX: A pedido de la gente, desactivo el comando /MATA
            Call WriteConsoleMsg(UserIndex, "El comando /MATA se encuentra desactivado.", FONTTYPE_SERVER)
            Exit Sub
        #End If
        
        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.Name, "/MATA " & Npclist(.flags.TargetNPC).Name)
        
    End With
    
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        
        #If ProteccionGM = 1 Then
            ' WyroX: A pedido de la gente, desactivo el comando /MASSKILL
            Call WriteConsoleMsg(UserIndex, "El comando /MASSKILL se encuentra desactivado.", FONTTYPE_SERVER)
            Exit Sub
        #End If
        
        Dim X       As Long
        
        Dim Y       As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.Map, X, Y).NpcIndex)
                    
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

Private Sub HandleLastIP(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName   As String
        
        Dim lista      As String
        
        Dim LoopC      As Byte
        
        Dim priv       As Integer
        
        Dim validCheck As Boolean
        
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.consejero
        UserName = Reader.ReadString8()
        
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
                
                If PersonajeExiste(UserName) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conecto son:" & vbCrLf & GetUserLastIps(UserName)
                    Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Else
                Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarquia que vos.", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Change the user`s chat color
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim color   As Long
        
        color = RGB(Reader.ReadInt8, Reader.ReadInt8, Reader.ReadInt8)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = color
            
        End If
        
    End With
    
End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Ignore the user
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
            
        End If
        
    End With
    
End Sub

''
' Handles the "CheckSlot" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 07/06/2010
    'Check one Users Slot in Particular from Inventory
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        'Reads the UserName and Slot Packets
        Dim UserName         As String
        
        Dim Slot             As Byte
        
        Dim tIndex           As Integer
        
        Dim UserIsAdmin      As Boolean
        
        Dim OtherUserIsAdmin As Boolean
        
        UserName = Reader.ReadString8()        'Que UserName
        Slot = Reader.ReadInt8        'Que Slot
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin Then
            
            Call LogGM(.Name, .Name & " Checkeo el slot " & Slot & " de " & UserName)
            
            tIndex = NameIndex(UserName)        'Que user index
            OtherUserIsAdmin = EsDios(UserName) Or EsAdmin(UserName)
            
            If tIndex > 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    If Slot > 0 And Slot <= UserList(tIndex).CurrentInventorySlots Then
                        If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
                            Call WriteConsoleMsg(UserIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).Name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(UserIndex, "No hay ningUn objeto en slot seleccionado.", FontTypeNames.FONTTYPE_INFO)
                            
                        End If
                        
                    Else
                        Call WriteConsoleMsg(UserIndex, "Slot Invalido.", FontTypeNames.FONTTYPE_TALK)
                        
                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver slots de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            Else
                
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver slots de un dios o admin.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "ResetAutoUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleResetAutoUpdate(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reset the AutoUpdate
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub
        
        Call WriteConsoleMsg(UserIndex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handles the "Restart" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleRestart(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Restart the game
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub
        
        'time and Time BUG!
        Call LogGM(.Name, .Name & " reinicio el mundo.")
        
        Call ReiniciarServidor(True)
        
    End With
    
End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the objects
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los objetos.")
        
        Call LoadOBJData
        
    End With
    
End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the spells
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los hechizos.")
        
        Call CargarHechizos
        
    End With
    
End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the Server`s INI
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los INITs.")
        
        Call LoadSini
        
        Call WriteConsoleMsg(UserIndex, "Server.ini actualizado correctamente", FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handle the "ReloadNPCs" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the Server`s NPC
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los NPCs.")
        
        Call CargaNpcsDat
        
        Call WriteConsoleMsg(UserIndex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Kick all the chars that are online
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados
        
    End With
    
End Sub

''
' Handle the "Night" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNight(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub
        
        DeNoche = Not DeNoche
        
        Dim i       As Long
        
        For i = 1 To NumUsers
            
            If UserList(i).flags.UserLogged And UserList(i).ConnID > -1 Then
                Call EnviarNoche(i)
                
            End If
            
        Next i
        
    End With
    
End Sub

''
' Handle the "ShowServerForm" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Show the server form
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click
        
    End With
    
End Sub

''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Clean the SOS
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha borrado los SOS.")
        
        Call Ayuda.Reset
        
    End With
    
End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Save the characters
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado todos los chars.")
        
        Call mdParty.ActualizaExperiencias
        Call GuardarUsuarios
        
    End With
    
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Change the backup`s info of the map
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim doTheBackUp As Boolean
        
        doTheBackUp = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la informacion sobre el BackUp.")
        
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.Map).BackUp = 1
        Else
            MapInfo(.Pos.Map).BackUp = 0
            
        End If
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo(.Pos.Map).BackUp)
        
        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).BackUp, FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Change the pk`s info of the  map
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim isMapPk As Boolean
        
        isMapPk = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si es PK el mapa.")
        
        MapInfo(.Pos.Map).Pk = isMapPk
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.Path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "Pk", IIf(isMapPk, "1", "0"))
        
        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " PK: " & MapInfo(.Pos.Map).Pk, FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
    '***************************************************
    
    On Error GoTo errHandler
    
    Dim tStr        As String
    
    With UserList(UserIndex)
        
        tStr = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si es restringido el mapa.")
                
                MapInfo(UserList(UserIndex).Pos.Map).Restringir = RestrictStringToByte(tStr)
                
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Restringir", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Restringido: " & RestrictByteToString(MapInfo(.Pos.Map).Restringir), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para restringir:        'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'MagiaSinEfecto -> Options: "1" , "0".
    '***************************************************
    
    Dim nomagic     As Boolean
    
    With UserList(UserIndex)
        
        nomagic = Reader.ReadBool
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar la magia el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto = nomagic
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " MagiaSinEfecto: " & MapInfo(.Pos.Map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handle the "ChangeMapInfoNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'InviSinEfecto -> Options: "1", "0"
    '***************************************************
    
    Dim noinvi      As Boolean
    
    With UserList(UserIndex)
        
        noinvi = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar la invisibilidad en el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).InviSinEfecto = noinvi
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " InviSinEfecto: " & MapInfo(.Pos.Map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handle the "ChangeMapInfoNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'ResuSinEfecto -> Options: "1", "0"
    '***************************************************
    
    Dim noresu      As Boolean
    
    With UserList(UserIndex)
        
        noresu = Reader.ReadBool()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar el resucitar en el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).ResuSinEfecto = noresu
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " ResuSinEfecto: " & MapInfo(.Pos.Map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    '***************************************************
    
    On Error GoTo errHandler
    
    Dim tStr        As String
    
    With UserList(UserIndex)
        
        tStr = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la informacion del terreno del mapa.")
                
                MapInfo(UserList(UserIndex).Pos.Map).Terreno = TerrainStringToByte(tStr)
                
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Terreno", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Terreno: " & TerrainByteToString(MapInfo(.Pos.Map).Terreno), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno:        'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el Unico Util es        'NIEVE' ya que al ingresarlo, la gente muere de frio en el mapa.", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    '***************************************************
    
    On Error GoTo errHandler
    
    Dim tStr        As String
    
    With UserList(UserIndex)
        
        tStr = Reader.ReadString8()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la informacion de la zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Zona = tStr
                Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno:        'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "        'DUNGEON', NO se sentira el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "        'NIEVE', Les agarra frio y saca salud hasta morir sin ropa.", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "ChangeMapInfoStealNp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoStealNpc(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 25/07/2010
    'RoboNpcsPermitido -> Options: "1", "0"
    '***************************************************
    
    Dim RoboNpc     As Byte
    
    With UserList(UserIndex)
        
        RoboNpc = val(IIf(Reader.ReadBool(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido robar npcs en el mapa.")
            
            MapInfo(UserList(UserIndex).Pos.Map).RoboNpcsPermitido = RoboNpc
            
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "RoboNpcsPermitido", RoboNpc)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " RoboNpcsPermitido: " & MapInfo(.Pos.Map).RoboNpcsPermitido, FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handle the "ChangeMapInfoNoOcultar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoOcultar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/09/2010
    'OcultarSinEfecto -> Options: "1", "0"
    '***************************************************
    
    Dim NoOcultar   As Byte
    
    Dim Mapa        As Integer
    
    With UserList(UserIndex)
        
        NoOcultar = val(IIf(Reader.ReadBool(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            
            Mapa = .Pos.Map
            
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido ocultarse en el mapa " & Mapa & ".")
            
            MapInfo(Mapa).OcultarSinEfecto = NoOcultar
            
            Call WriteVar(App.Path & MapPath & "mapa" & Mapa & ".dat", "Mapa" & Mapa, "OcultarSinEfecto", NoOcultar)
            Call WriteConsoleMsg(UserIndex, "Mapa " & Mapa & " OcultarSinEfecto: " & NoOcultar, FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handle the "ChangeMapInfoNoInvocar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoInvocar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/09/2010
    'InvocarSinEfecto -> Options: "1", "0"
    '***************************************************
    
    Dim NoInvocar   As Byte
    
    Dim Mapa        As Integer
    
    With UserList(UserIndex)
        
        NoInvocar = val(IIf(Reader.ReadBool(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            
            Mapa = .Pos.Map
            
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido invocar en el mapa " & Mapa & ".")
            
            MapInfo(Mapa).InvocarSinEfecto = NoInvocar
            
            Call WriteVar(App.Path & MapPath & "mapa" & Mapa & ".dat", "Mapa" & Mapa, "InvocarSinEfecto", NoInvocar)
            Call WriteConsoleMsg(UserIndex, "Mapa " & Mapa & " InvocarSinEfecto: " & NoInvocar, FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Saves the map
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.Map))
        
        Call GrabarMapa(.Pos.Map, App.Path & "\WorldBackUp\Mapa" & .Pos.Map)
        
        Call WriteConsoleMsg(UserIndex, "Mapa Guardado.", FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Allows admins to read guild messages
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Guild   As String
        
        Guild = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(UserIndex, Guild)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Show guilds messages
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha hecho un backup.")
        
        Call ES.DoBackUp        'Sino lo confunde con la id del paquete
        
    End With
    
End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 02/05/2012
    'Nuevo centinela (maTih.-)
    '***************************************************
    With UserList(UserIndex)
        
        'Solo para Admins y Dioses
        If Not EsAdmin(.Name) Or Not EsDios(.Name) Then Exit Sub
        
        Call modCentinela.CambiarEstado(UserIndex)
        
    End With
    
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Change user name
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        'Reads the userName and newUser Packets
        Dim UserName     As String
        
        Dim newName      As String
        
        Dim changeNameUI As Integer
        
        Dim GuildIndex   As Integer
        
        UserName = Reader.ReadString8()
        newName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El Pj esta online, debe salir para hacer el cambio.", FontTypeNames.FONTTYPE_WARNING)
                Else
                    
                    If Not PersonajeExiste(UserName) Then
                        Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " es inexistente.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        
                        If GetUserGuildIndex(UserName) > 0 Then
                            Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            
                            If Not PersonajeExiste(newName) Then
                                Call CopyUser(UserName, newName)
                                
                                If Not Database_Enabled Then
                                    Call SaveBan(UserName, "BAN POR Cambio de nick a " & UCase$(newName), .Name)
                                    
                                End If
                                
                                Call WriteConsoleMsg(UserIndex, "Transferencia exitosa.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.Name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)
                                
                            End If
                            
                        End If
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterMail(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Change user password
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim newMail  As String
        
        UserName = Reader.ReadString8()
        newMail = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
            Else
                
                If Not PersonajeExiste(UserName) Then
                    Call WriteConsoleMsg(UserIndex, "No existe el charfile de" & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SaveUserEmail(UserName, newMail)
                    Call WriteConsoleMsg(UserIndex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
                Call LogGM(.Name, "Le ha cambiado el mail a " & UserName)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "AlterPassword" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Change user password
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim copyFrom As String
        
        Dim Password As String
        
        Dim Salt     As String
        
        UserName = Replace(Reader.ReadString8(), "+", " ")
        copyFrom = Replace(Reader.ReadString8(), "+", " ")
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha alterado la contrasena de " & UserName)
            
            If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /APASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
            Else
                
                If Not PersonajeExiste(UserName) Or Not PersonajeExiste(copyFrom) Then
                    Call WriteConsoleMsg(UserIndex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
                Else
                    Password = GetUserPassword(copyFrom)
                    Salt = GetUserSalt(copyFrom)
                    
                    Call StorePasswordSalt(UserName, Password, Salt)
                    
                    Call WriteConsoleMsg(UserIndex, "Password de " & UserName & " ha cambiado por la de " & copyFrom, FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)
    
    '**********************************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/05/2019
    '26/09/2010: ZaMa - Ya no se pueden crear npcs pretorianos.
    '11/05/2019: Jopi - Se arreglo la comprobacion de NPC's pretorianos.
    '11/05/2019: Jopi - Se combino HandleCreateNPCWithRespawn() con este procedimiento.
    '**********************************************************************************
    
    With UserList(UserIndex)
        
        Dim NpcIndex As Integer: NpcIndex = Reader.ReadInt16()
        Dim Respawn As Boolean: Respawn = Reader.ReadBool()
        
        'Nos fijamos que sea GM.
        If Not EsGm(UserIndex) Then Exit Sub
        
        'Nos fijamos si es pretoriano.
        If Npclist(NpcIndex).NPCtype = eNPCType.Pretoriano Then
            Call WriteConsoleMsg(UserIndex, "No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CREARPRETORIANOS MAPA X Y.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
            
        End If
        
        'Invocamos el NPC.
        If NpcIndex <> 0 Then
            
            NpcIndex = SpawnNpc(NpcIndex, .Pos, True, Respawn)
            
            Call LogGM(.Name, "Invoco " & IIf(Respawn, "con respawn", vbNullString) & " a " & Npclist(NpcIndex).Name & " [Indice: " & NpcIndex & "] en el mapa " & .Pos.Map)
            
        End If
        
    End With
    
End Sub

''
' Handle the "ImperialArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleImperialArmour(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim index    As Byte
        
        Dim ObjIndex As Integer
        
        index = Reader.ReadInt8
        ObjIndex = Reader.ReadInt16()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
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

Public Sub HandleChaosArmour(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim index    As Byte
        
        Dim ObjIndex As Integer
        
        index = Reader.ReadInt8
        ObjIndex = Reader.ReadInt16()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
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

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/12/07
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero) Then Exit Sub
        
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1
            
        End If
        
        'Tell the client that we are navigating.
        Call WriteNavigateToggle(UserIndex)
        
    End With
    
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
            frmMain.chkServerHabilitado.Value = vbUnchecked
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
            frmMain.chkServerHabilitado.Value = vbChecked
            
        End If
        
    End With
    
End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    'Turns off the server
    '***************************************************
    Dim handle      As Integer
    
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "/APAGAR")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("" & .Name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
        
        'Log
        handle = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #handle
        
        Print #handle, Date & " " & Time & " server apagado por " & .Name & ". "
        
        Close #handle
        
        Unload frmMain
        
    End With
    
End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/CONDEN " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then Call VolverCriminal(tUser)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 06/09/09
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim tUser    As Integer
        
        Dim Char     As String
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            Call LogGM(.Name, "/RAJAR " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call ResetFacciones(tUser)
            Else
                
                If PersonajeExiste(UserName) Then
                    Call ResetUserFacciones(UserName)
                Else
                    Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName   As String
        
        Dim GuildIndex As Integer
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJARCLAN " & UserName)
            
            GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "No pertenece a ningUn clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "RequestCharMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Request user mail
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim mail     As String
        
        UserName = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            If PersonajeExiste(UserName) Then
                mail = GetUserEmail(UserName)
                
                Call WriteConsoleMsg(UserIndex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/29/06
    'Send a message to all the users
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Message As String
        
        Message = Reader.ReadString8()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Mensaje de sistema:" & Message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(Message))
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "SetMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetMOTD(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 03/31/07
    'Set the MOTD
    'Modified by: Juan Martin Sotuyo Dodero (Maraxus)
    '   - Fixed a bug that prevented from properly setting the new number of lines.
    '   - Fixed a bug that caused the player to be kicked.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim newMOTD           As String
        
        Dim auxiliaryString() As String
        
        Dim LoopC             As Long
        
        newMOTD = Reader.ReadString8()
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
            
            Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con exito.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "ChangeMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin sotuyo Dodero (Maraxus)
    'Last Modification: 12/29/06
    'Change the MOTD
    '***************************************************
    With UserList(UserIndex)
        
        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.User Or PlayerType.consejero Or PlayerType.SemiDios)) Then
            Exit Sub
            
        End If
        
        Dim auxiliaryString As String
        
        Dim LoopC           As Long
        
        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
        Next LoopC
        
        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
                
            End If
            
        End If
        
        Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)
        
    End With
    
End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Show guilds messages
    '***************************************************
    With UserList(UserIndex)
        
        Call WritePong(UserIndex)
        
    End With
    
End Sub

''
' Handle the "SetIniVar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetIniVar(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Brian Chaia (BrianPr)
    'Last Modification: 01/23/10 (Marco)
    'Modify server.ini
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim sLlave  As String
        
        Dim sClave  As String
        
        Dim sValor  As String
        
        'Obtengo los parametros
        sLlave = Reader.ReadString8()
        sClave = Reader.ReadString8()
        sValor = Reader.ReadString8()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            
            Dim sTmp As String
            
            'No podemos modificar [INIT]Dioses ni [Dioses]*
            If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "DIOSES") Or UCase$(sLlave) = "DIOSES" Then
                Call WriteConsoleMsg(UserIndex, "No puedes modificar esa informacion desde aqui!", FontTypeNames.FONTTYPE_INFO)
            Else
                'Obtengo el valor segUn llave y clave
                sTmp = GetVar(IniPath & "Server.ini", sLlave, sClave)
                
                'Si obtengo un valor escribo en el server.ini
                If LenB(sTmp) Then
                    Call WriteVar(IniPath & "Server.ini", sLlave, sClave, sValor)
                    Call LogGM(.Name, "Modifico en server.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & " por " & sValor)
                    Call WriteConsoleMsg(UserIndex, "Modifico " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)
                    
                End If
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreatePretorianClan(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    '***************************************************
    
    On Error GoTo errHandler
    
    Dim Map         As Integer
    Dim X           As Byte
    Dim Y           As Byte
    Dim index       As Long
    
    With UserList(UserIndex)
        
        Map = Reader.ReadInt16()
        X = Reader.ReadInt8
        Y = Reader.ReadInt8
        
        ' User Admin
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) = 0 Then Exit Sub
        
        ' Valid pos
        If Not InMapBounds(Map, X, Y) Then
            Call WriteConsoleMsg(UserIndex, "Posicion invalida.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        ' Choose pretorian clan index
        If Map = MAPA_PRETORIANO Then
            index = ePretorianType.Default        ' Default clan
        Else
            index = ePretorianType.Custom        ' Custom Clan
        End If
        
        ' Is already active any clan
        If Not ClanPretoriano(index).Active Then
            
            If Not ClanPretoriano(index).SpawnClan(Map, X, Y, index) Then
                Call WriteConsoleMsg(UserIndex, "La posicion no es apropiada para crear el clan", FontTypeNames.FONTTYPE_INFO)
                
            End If
            
        Else
            Call WriteConsoleMsg(UserIndex, "El clan pretoriano se encuentra activo en el mapa " & ClanPretoriano(index).ClanMap & ". Utilice /EliminarPretorianos MAPA y reintente.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With
    
    Exit Sub
    
errHandler:
    Call LogError("Error en HandleCreatePretorianClan. Error: " & Err.Number & " - " & Err.description)
    
End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDeletePretorianClan(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    '***************************************************
    
    On Error GoTo errHandler
    
    Dim Map         As Integer
    
    Dim index       As Long
    
    With UserList(UserIndex)
        
        Map = Reader.ReadInt16()
        
        ' User Admin
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) = 0 Then Exit Sub
        
        ' Valid map
        If Map < 1 Or Map > NumMaps Then
            Call WriteConsoleMsg(UserIndex, "Mapa invalido.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            
        End If
        
        For index = 1 To UBound(ClanPretoriano)
            
            ' Search for the clan to be deleted
            If ClanPretoriano(index).ClanMap = Map Then
                ClanPretoriano(index).DeleteClan
                Exit For
                
            End If
            
        Next index
        
    End With
    
    Exit Sub
    
errHandler:
    Call LogError("Error en HandleDeletePretorianClan. Error: " & Err.Number & " - " & Err.description)
    
End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Logged" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        Call Writer.WriteInt8(ServerPacketID.Logged)
        #If AntiExternos Then
            UserList(UserIndex).Redundance = RandomNumber(15, 250)
            Call Writer.WriteInt8(UserList(UserIndex).Redundance)
        #End If
        Call Writer.WriteInt8(.Clase)
        Call Writer.WriteInt32(IntervaloInvisible)
    End With
    
    Exit Sub
    
errHandler:
    
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.RemoveDialogs)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageRemoveCharDialog(CharIndex))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.NavigateToggle)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Disconnect" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.Disconnect)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UserOfferConfirm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserOfferConfirm(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    'Writes the "UserOfferConfirm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UserOfferConfirm)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.CommerceEnd)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.BankEnd)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.CommerceInit)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.BankInit)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Banco)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UserCommerceInit)
    Call Writer.WriteString8(UserList(UserIndex).ComUsu.DestNick)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UserCommerceEnd)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UpdateSta)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UpdateMana)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMAN)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UpdateHP)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateGold" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UpdateGold)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Gld)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UpdateBankGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateBankGold(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    'Writes the "UpdateBankGold" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UpdateBankGold)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Banco)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateExp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UpdateExp)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenghtAndDexterity(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UpdateStrenghtAndDexterity)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    Call Writer.WriteInt32(UserList(UserIndex).flags.DuracionEfecto)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDexterity(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UpdateDexterity)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    Call Writer.WriteInt32(UserList(UserIndex).flags.DuracionEfecto)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenght(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UpdateStrenght)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
    Call Writer.WriteInt32(UserList(UserIndex).flags.DuracionEfecto)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, _
       ByVal Map As Integer, _
       ByVal version As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMap" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ChangeMap)
    Call Writer.WriteInt16(Map)
    Call Writer.WriteString8(MapInfo(Map).Name)
    Call Writer.WriteString8(MapInfo(Map).Zona)
    Call Writer.WriteInt16(version)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PosUpdate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.PosUpdate)
    Call Writer.WriteInt8(UserList(UserIndex).Pos.X)
    Call Writer.WriteInt8(UserList(UserIndex).Pos.Y)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, _
       ByVal Chat As String, _
       ByVal CharIndex As Integer, _
       ByVal color As Long)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageChatOverHead(Chat, CharIndex, color))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, _
       ByVal Chat As String, _
       ByVal FontIndex As FontTypeNames)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageConsoleMsg(Chat, FontIndex))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub
Public Sub WriteRenderMsg(ByVal UserIndex As Integer, _
       ByVal Chat As String, _
       ByVal FontIndex As Integer)
    
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareRenderConsoleMsg(Chat, FontIndex))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

Public Sub WriteCommerceChat(ByVal UserIndex As Integer, _
       ByVal Chat As String, _
       ByVal FontIndex As FontTypeNames)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareCommerceConsoleMsg(Chat, FontIndex))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal Chat As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildChat" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageGuildChat(Chat))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal Message As String)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
    '***************************************************
    
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ShowMessageBox)
    Call Writer.WriteString8(Message)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UserIndexInServer)
    Call Writer.WriteInt16(UserIndex)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UserCharIndexInServer)
    Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
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

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, _
       ByVal body As Integer, _
       ByVal Head As Integer, _
       ByVal heading As eHeading, _
       ByVal CharIndex As Integer, _
       ByVal X As Byte, _
       ByVal Y As Byte, _
       ByVal weapon As Integer, _
       ByVal shield As Integer, _
       ByVal FX As Integer, _
       ByVal FXLoops As Integer, _
       ByVal helmet As Integer, _
       ByVal Name As String, _
       ByVal NickColor As Byte, _
       ByVal Privileges As Byte)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterCreate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageCharacterCreate(body, Head, heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, helmet, Name, NickColor, Privileges))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterRemove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageCharacterRemove(CharIndex))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, _
       ByVal CharIndex As Integer, _
       ByVal X As Byte, _
       ByVal Y As Byte)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageCharacterMove(CharIndex, X, Y))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Writes the "ForceCharMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageForceCharMove(Direccion))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
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

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, _
       ByVal body As Integer, _
       ByVal Head As Integer, _
       ByVal heading As eHeading, _
       ByVal CharIndex As Integer, _
       ByVal weapon As Integer, _
       ByVal shield As Integer, _
       ByVal FX As Integer, _
       ByVal FXLoops As Integer, _
       ByVal helmet As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterChange" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageCharacterChange(body, Head, heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, _
       ByVal GrhIndex As Long, _
       ByVal X As Byte, _
       ByVal Y As Byte)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectCreate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageObjectCreate(GrhIndex, X, Y))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectDelete" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageObjectDelete(X, Y))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, _
       ByVal X As Byte, _
       ByVal Y As Byte, _
       ByVal Blocked As Boolean)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockPosition" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.BlockPosition)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
    Call Writer.WriteBool(Blocked)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "PlayMp3" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    mp3 The mp3 to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMp3(ByVal UserIndex As Integer, _
       ByVal mp3 As Integer, _
       Optional ByVal loops As Integer = -1)
    
    '***************************************************
    'Author: Lucas Recoaro (Recox)
    'Last Modification: 05/17/06
    'Writes the "PlayMp3" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessagePlayMp3(mp3, loops))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal UserIndex As Integer, _
       ByVal midi As Integer, _
       Optional ByVal loops As Integer = -1)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PlayMidi" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessagePlayMidi(midi, loops))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal UserIndex As Integer, _
       ByVal wave As Byte, _
       ByVal X As Byte, _
       ByVal Y As Byte)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessagePlayWave(wave, X, Y))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim Tmp         As String
    
    Dim i           As Long
    
    Call Writer.WriteInt8(ServerPacketID.guildList)
    
    ' Prepare guild name's list
    For i = LBound(guildList()) To UBound(guildList())
        Tmp = Tmp & guildList(i) & SEPARATOR
    Next i
    
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AreaChanged" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.AreaChanged)
    Call Writer.WriteInt8(UserList(UserIndex).Pos.X)
    Call Writer.WriteInt8(UserList(UserIndex).Pos.Y)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PauseToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessagePauseToggle())
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageRainToggle())
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, _
       ByVal CharIndex As Integer, _
       ByVal FX As Integer, _
       ByVal FXLoops As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateFX" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UpdateUserStats)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxHp)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxMAN)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMAN)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxSta)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Gld)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.ELV)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.ELU)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 3/12/09
    'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
    '3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ChangeInventorySlot)
    Call Writer.WriteInt8(Slot)
    
    Dim ObjIndex    As Integer
    
    Dim obData      As ObjData
    
    ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    
    If ObjIndex > 0 Then
        obData = ObjData(ObjIndex)
        
    End If
    
    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteString8(obData.Name)
    Call Writer.WriteInt16(UserList(UserIndex).Invent.Object(Slot).Amount)
    Call Writer.WriteBool(UserList(UserIndex).Invent.Object(Slot).Equipped)
    Call Writer.WriteInt32(obData.GrhIndex)
    Call Writer.WriteInt8(obData.OBJType)
    Call Writer.WriteInt16(obData.MaxHIT)
    Call Writer.WriteInt16(obData.MinHIT)
    Call Writer.WriteInt16(obData.MaxDef)
    Call Writer.WriteInt16(obData.MinDef)
    Call Writer.WriteReal32(SalePrice(ObjIndex))
    Call Writer.WriteBool(ItemIncompatibleConUser(UserIndex, ObjIndex))
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

Public Sub WriteAddSlots(ByVal UserIndex As Integer, ByVal Mochila As eMochilas)
    
    '***************************************************
    'Author: Budi
    'Last Modification: 01/12/09
    'Writes the "AddSlots" message to the given user's outgoing data buffer
    '***************************************************
    Call Writer.WriteInt8(ServerPacketID.AddSlots)
    Call Writer.WriteInt8(Mochila)
    
End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de solo Def
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ChangeBankSlot)
    Call Writer.WriteInt8(Slot)
    
    Dim ObjIndex    As Integer
    
    Dim obData      As ObjData
    
    ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
    
    Call Writer.WriteInt16(ObjIndex)
    
    If ObjIndex > 0 Then
        obData = ObjData(ObjIndex)
        
    End If
    
    Call Writer.WriteString8(obData.Name)
    Call Writer.WriteInt16(UserList(UserIndex).BancoInvent.Object(Slot).Amount)
    Call Writer.WriteInt32(obData.GrhIndex)
    Call Writer.WriteInt8(obData.OBJType)
    Call Writer.WriteInt16(obData.MaxHIT)
    Call Writer.WriteInt16(obData.MinHIT)
    Call Writer.WriteInt16(obData.MaxDef)
    Call Writer.WriteInt16(obData.MinDef)
    Call Writer.WriteInt32(obData.Valor)
    Call Writer.WriteBool(ItemIncompatibleConUser(UserIndex, ObjIndex))
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 27/08/2016
    'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
    '27-08-2016: Shak@ Gracias a la optimizacion, enviamos menos datos :P
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ChangeSpellSlot)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.UserHechizos(Slot))
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Atributes" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.Atributes)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
    'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i              As Long
    Dim obj            As ObjData
    Dim validIndexes() As Integer
    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    Call Writer.WriteInt8(ServerPacketID.BlacksmithWeapons)
    
    For i = 1 To UBound(ArmasHerrero())
        
        ' Can the user create this object? If so add it to the list....
        If ObjData(ArmasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).Clase), 0) Then
            Count = Count + 1
            validIndexes(Count) = i
            
        End If
        
    Next i
    
    ' Write the number of objects in the list
    Call Writer.WriteInt16(Count)
    
    ' Write the needed data of each object
    For i = 1 To Count
        obj = ObjData(ArmasHerrero(validIndexes(i)))
        Call Writer.WriteString8(obj.Name)
        Call Writer.WriteInt32(obj.GrhIndex)
        Call Writer.WriteInt16(obj.LingH)
        Call Writer.WriteInt16(obj.LingP)
        Call Writer.WriteInt16(obj.LingO)
        Call Writer.WriteInt16(ArmasHerrero(validIndexes(i)))
        Call Writer.WriteInt16(obj.Upgrade)
    Next i
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
    'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i              As Long
    Dim obj            As ObjData
    Dim validIndexes() As Integer
    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    Call Writer.WriteInt8(ServerPacketID.BlacksmithArmors)
    
    For i = 1 To UBound(ArmadurasHerrero())
        
        ' Can the user create this object? If so add it to the list....
        If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).Clase), 0) Then
            Count = Count + 1
            validIndexes(Count) = i
            
        End If
        
    Next i
    
    ' Write the number of objects in the list
    Call Writer.WriteInt16(Count)
    
    ' Write the needed data of each object
    For i = 1 To Count
        obj = ObjData(ArmadurasHerrero(validIndexes(i)))
        Call Writer.WriteString8(obj.Name)
        Call Writer.WriteInt32(obj.GrhIndex)
        Call Writer.WriteInt16(obj.LingH)
        Call Writer.WriteInt16(obj.LingP)
        Call Writer.WriteInt16(obj.LingO)
        Call Writer.WriteInt16(ArmadurasHerrero(validIndexes(i)))
        Call Writer.WriteInt16(obj.Upgrade)
    Next i
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "InitCarpenting" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInitCarpenting(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "InitCarpenting" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i              As Long
    Dim obj            As ObjData
    Dim validIndexes() As Integer
    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    
    Call Writer.WriteInt8(ServerPacketID.InitCarpenting)
    
    For i = 1 To UBound(ObjCarpintero())
        
        ' Can the user create this object? If so add it to the list....
        If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).Clase) Then
            Count = Count + 1
            validIndexes(Count) = i
            
        End If
        
    Next i
    
    ' Write the number of objects in the list
    Call Writer.WriteInt16(Count)
    
    ' Write the needed data of each object
    For i = 1 To Count
        obj = ObjData(ObjCarpintero(validIndexes(i)))
        Call Writer.WriteString8(obj.Name)
        Call Writer.WriteInt32(obj.GrhIndex)
        Call Writer.WriteInt16(obj.Madera)
        Call Writer.WriteInt16(obj.MaderaElfica)
        Call Writer.WriteInt16(ObjCarpintero(validIndexes(i)))
        Call Writer.WriteInt16(obj.Upgrade)
    Next i
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

Public Sub WriteInitCraftsman(ByVal UserIndex As Integer)
    '***************************************************
    'Author: WyroX
    'Last Modification: 27/01/2020
    'Writes the "InitCraftman" message to the given user's outgoing data buffer
    '***************************************************
    
    On Error GoTo errHandler
    
    Dim i              As Long
    Dim j              As Long
    Dim obj            As ObjData
    Dim ObjRequired    As ObjData
    
    Call Writer.WriteInt8(ServerPacketID.InitCraftman)
    
    ' Write cost of crafting
    Call Writer.WriteInt32(ArtesaniaCosto)
    
    ' Write the number of objects in the list
    Call Writer.WriteInt16(UBound(ObjArtesano))
    
    ' Write the needed data of each object
    For i = 1 To UBound(ObjArtesano)
        obj = ObjData(ObjArtesano(i))
        Call Writer.WriteString8(obj.Name)
        Call Writer.WriteInt32(obj.GrhIndex)
        Call Writer.WriteInt16(ObjArtesano(i))
        
        Call Writer.WriteInt8(UBound(obj.ItemCrafteo))
        
        For j = 1 To UBound(obj.ItemCrafteo)
            ObjRequired = ObjData(obj.ItemCrafteo(j).ObjIndex)
            Call Writer.WriteString8(ObjRequired.Name)
            Call Writer.WriteInt32(ObjRequired.GrhIndex)
            Call Writer.WriteInt16(obj.ItemCrafteo(j).ObjIndex)
            Call Writer.WriteInt16(obj.ItemCrafteo(j).Amount)
        Next j
    Next i
    
    Exit Sub
    
errHandler:
    
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RestOK" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.RestOK)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal Message As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ErrorMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageErrorMsg(Message))
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Blind" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.Blind)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Dumb" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.Dumb)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowSignal" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ShowSignal)
    Call Writer.WriteString8(ObjData(ObjIndex).texto)
    Call Writer.WriteInt32(ObjData(ObjIndex).GrhSecundario)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, _
       ByVal Slot As Byte, _
       ByRef obj As obj, _
       ByVal price As Single)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Last Modified by: Budi
    'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de solo Def
    '***************************************************
    On Error GoTo errHandler
    
    Dim ObjInfo     As ObjData
    
    If obj.ObjIndex >= LBound(ObjData()) And obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(obj.ObjIndex)
        
    End If
    
    Call Writer.WriteInt8(ServerPacketID.ChangeNPCInventorySlot)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteString8(ObjInfo.Name)
    Call Writer.WriteInt16(obj.Amount)
    Call Writer.WriteReal32(price)
    Call Writer.WriteInt32(ObjInfo.GrhIndex)
    Call Writer.WriteInt16(obj.ObjIndex)
    Call Writer.WriteInt8(ObjInfo.OBJType)
    Call Writer.WriteInt16(ObjInfo.MaxHIT)
    Call Writer.WriteInt16(ObjInfo.MinHIT)
    Call Writer.WriteInt16(ObjInfo.MaxDef)
    Call Writer.WriteInt16(ObjInfo.MinDef)
    Call Writer.WriteBool(ItemIncompatibleConUser(UserIndex, obj.ObjIndex))
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UpdateHungerAndThirst)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.MaxAGU)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.MinAGU)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.MaxHam)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.MinHam)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "Fame" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteFame(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.Fame)
    
    Call Writer.WriteInt32(UserList(UserIndex).Reputacion.AsesinoRep)
    Call Writer.WriteInt32(UserList(UserIndex).Reputacion.BandidoRep)
    Call Writer.WriteInt32(UserList(UserIndex).Reputacion.BurguesRep)
    Call Writer.WriteInt32(UserList(UserIndex).Reputacion.LadronesRep)
    Call Writer.WriteInt32(UserList(UserIndex).Reputacion.NobleRep)
    Call Writer.WriteInt32(UserList(UserIndex).Reputacion.PlebeRep)
    Call Writer.WriteInt32(UserList(UserIndex).Reputacion.Promedio)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MiniStats" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.MiniStats)
    
    Call Writer.WriteInt32(UserList(UserIndex).Faccion.CiudadanosMatados)
    Call Writer.WriteInt32(UserList(UserIndex).Faccion.CriminalesMatados)
    
    'TODO : Este valor es calculable, no deberia NI EXISTIR, ya sea en el servidor ni en el cliente!!!
    Call Writer.WriteInt32(UserList(UserIndex).Stats.UsuariosMatados)
    
    Call Writer.WriteInt16(UserList(UserIndex).Stats.NPCsMuertos)
    
    Call Writer.WriteInt8(UserList(UserIndex).Clase)
    Call Writer.WriteInt32(UserList(UserIndex).Counters.Pena)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LevelUp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.LevelUp)
    Call Writer.WriteInt16(skillPoints)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, _
       ByVal ForumType As eForumType, _
       ByRef Title As String, _
       ByRef Author As String, _
       ByRef Message As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 02/01/2010
    'Writes the "AddForumMsg" message to the given user's outgoing data buffer
    '02/01/2010: ZaMa - Now sends Author and forum type
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.AddForumMsg)
    Call Writer.WriteInt8(ForumType)
    Call Writer.WriteString8(Title)
    Call Writer.WriteString8(Author)
    Call Writer.WriteString8(Message)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowForumForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim Visibilidad   As Byte
    
    Dim CanMakeSticky As Byte
    
    With UserList(UserIndex)
        Call Writer.WriteInt8(ServerPacketID.ShowForumForm)
        
        Visibilidad = eForumVisibility.ieGENERAL_MEMBER
        
        If esCaos(UserIndex) Or EsGm(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieCAOS_MEMBER
            
        End If
        
        If esArmada(UserIndex) Or EsGm(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieREAL_MEMBER
            
        End If
        
        Call Writer.WriteInt8(Visibilidad)
        
        ' Pueden mandar sticky los gms o los del consejo de armada/caos
        If EsGm(UserIndex) Then
            CanMakeSticky = 2
        ElseIf (.flags.Privilegios And PlayerType.ChaosCouncil) <> 0 Then
            CanMakeSticky = 1
        ElseIf (.flags.Privilegios And PlayerType.RoyalCouncil) <> 0 Then
            CanMakeSticky = 1
            
        End If
        
        Call Writer.WriteInt8(CanMakeSticky)
        
    End With
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, _
       ByVal CharIndex As Integer, _
       ByVal invisible As Boolean)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetInvisible" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteString8(PrepareMessageSetInvisible(CharIndex, invisible))
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "DiceRoll" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDiceRoll(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DiceRoll" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.DiceRoll)
    
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MeditateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.MeditateToggle)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlindNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.BlindNoMore)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DumbNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.DumbNoMore)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/19/09
    'Writes the "SendSkills" message to the given user's outgoing data buffer
    '11/19/09: Pato - Now send the percentage of progress of the skills.
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    With UserList(UserIndex)
        Call Writer.WriteInt8(ServerPacketID.SendSkills)
        Call Writer.WriteInt8(.Clase)
        
        For i = 1 To NUMSKILLS
            Call Writer.WriteInt8(UserList(UserIndex).Stats.UserSkills(i))
            
            If .Stats.UserSkills(i) < MAXSKILLPOINTS Then
                Call Writer.WriteInt8(Int(.Stats.ExpSkills(i) * 100 / .Stats.EluSkills(i)))
            Else
                Call Writer.WriteInt8(0)
                
            End If
            
        Next i
        
    End With
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    Dim Str         As String
    
    Call Writer.WriteInt8(ServerPacketID.TrainerCreatureList)
    
    For i = 1 To Npclist(NpcIndex).NroCriaturas
        Str = Str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
    Next i
    
    If LenB(Str) > 0 Then Str = Left$(Str, Len(Str) - 1)
    
    Call Writer.WriteString8(Str)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal UserIndex As Integer, _
       ByVal guildNews As String, _
       ByRef enemies() As String, _
       ByRef allies() As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildNews" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    Dim Tmp         As String
    
    Call Writer.WriteInt8(ServerPacketID.guildNews)
    
    Call Writer.WriteString8(guildNews)
    
    'Prepare enemies' list
    For i = LBound(enemies()) To UBound(enemies())
        Tmp = Tmp & enemies(i) & SEPARATOR
    Next i
    
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    Tmp = vbNullString
    
    'Prepare allies' list
    For i = LBound(allies()) To UBound(allies())
        Tmp = Tmp & allies(i) & SEPARATOR
    Next i
    
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OfferDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    Call Writer.WriteInt8(ServerPacketID.OfferDetails)
    
    Call Writer.WriteString8(details)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    Dim Tmp         As String
    
    Call Writer.WriteInt8(ServerPacketID.AlianceProposalsList)
    
    ' Prepare guild's list
    For i = LBound(guilds()) To UBound(guilds())
        Tmp = Tmp & guilds(i) & SEPARATOR
    Next i
    
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    Dim Tmp         As String
    
    Call Writer.WriteInt8(ServerPacketID.PeaceProposalsList)
    
    ' Prepare guilds' list
    For i = LBound(guilds()) To UBound(guilds())
        Tmp = Tmp & guilds(i) & SEPARATOR
    Next i
    
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
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

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal charName As String, ByVal race As eRaza, ByVal Class As eClass, ByVal gender As eGenero, ByVal level As Byte, ByVal Gold As Long, ByVal bank As Long, ByVal reputation As Long, ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.CharacterInfo)
    
    Call Writer.WriteString8(charName)
    Call Writer.WriteInt8(race)
    Call Writer.WriteInt8(Class)
    Call Writer.WriteInt8(gender)
    
    Call Writer.WriteInt8(level)
    Call Writer.WriteInt32(Gold)
    Call Writer.WriteInt32(bank)
    Call Writer.WriteInt32(reputation)
    
    Call Writer.WriteString8(previousPetitions)
    Call Writer.WriteString8(currentGuild)
    Call Writer.WriteString8(previousGuilds)
    
    Call Writer.WriteBool(RoyalArmy)
    Call Writer.WriteBool(CaosLegion)
    
    Call Writer.WriteInt32(citicensKilled)
    Call Writer.WriteInt32(criminalsKilled)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
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

Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, _
       ByRef guildList() As String, _
       ByRef MemberList() As String, _
       ByVal guildNews As String, _
       ByRef joinRequests() As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    Dim Tmp         As String
    
    Call Writer.WriteInt8(ServerPacketID.GuildLeaderInfo)
    
    ' Prepare guild name's list
    For i = LBound(guildList()) To UBound(guildList())
        Tmp = Tmp & guildList(i) & SEPARATOR
    Next i
    
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    ' Prepare guild member's list
    Tmp = vbNullString
    
    For i = LBound(MemberList()) To UBound(MemberList())
        Tmp = Tmp & MemberList(i) & SEPARATOR
    Next i
    
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    ' Store guild news
    Call Writer.WriteString8(guildNews)
    
    ' Prepare the join request's list
    Tmp = vbNullString
    
    For i = LBound(joinRequests()) To UBound(joinRequests())
        Tmp = Tmp & joinRequests(i) & SEPARATOR
    Next i
    
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserIndex As Integer, _
       ByRef guildList() As String, _
       ByRef MemberList() As String)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 21/02/2010
    'Writes the "GuildMemberInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    Dim Tmp         As String
    
    Call Writer.WriteInt8(ServerPacketID.GuildMemberInfo)
    
    ' Prepare guild name's list
    For i = LBound(guildList()) To UBound(guildList())
        Tmp = Tmp & guildList(i) & SEPARATOR
    Next i
    
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    ' Prepare guild member's list
    Tmp = vbNullString
    
    For i = LBound(MemberList()) To UBound(MemberList())
        Tmp = Tmp & MemberList(i) & SEPARATOR
    Next i
    
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
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

Public Sub WriteGuildDetails(ByVal UserIndex As Integer, _
       ByVal GuildName As String, _
       ByVal founder As String, _
       ByVal foundationDate As String, _
       ByVal leader As String, _
       ByVal URL As String, _
       ByVal memberCount As Integer, _
       ByVal electionsOpen As Boolean, _
       ByVal alignment As String, _
       ByVal enemiesCount As Integer, _
       ByVal AlliesCount As Integer, _
       ByVal antifactionPoints As String, _
       ByRef codex() As String, _
       ByVal guildDesc As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    Dim temp        As String
    
    Call Writer.WriteInt8(ServerPacketID.GuildDetails)
    
    Call Writer.WriteString8(GuildName)
    Call Writer.WriteString8(founder)
    Call Writer.WriteString8(foundationDate)
    Call Writer.WriteString8(leader)
    Call Writer.WriteString8(URL)
    
    Call Writer.WriteInt16(memberCount)
    Call Writer.WriteBool(electionsOpen)
    
    Call Writer.WriteString8(alignment)
    
    Call Writer.WriteInt16(enemiesCount)
    Call Writer.WriteInt16(AlliesCount)
    
    Call Writer.WriteString8(antifactionPoints)
    
    For i = LBound(codex()) To UBound(codex())
        temp = temp & codex(i) & SEPARATOR
    Next i
    
    If Len(temp) > 1 Then temp = Left$(temp, Len(temp) - 1)
    
    Call Writer.WriteString8(temp)
    
    Call Writer.WriteString8(guildDesc)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ShowGuildAlign" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildAlign(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    'Writes the "ShowGuildAlign" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ShowGuildAlign)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ShowGuildFundationForm)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 08/12/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    'Writes the "ParalizeOK" message to the given user's outgoing data buffer
    'And updates user position
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ParalizeOK)
    Call Writer.WriteInt16(IIf(UserList(UserIndex).flags.Paralizado, UserList(UserIndex).Counters.Paralisis, 0))
    
    Call WritePosUpdate(UserIndex)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ShowUserRequest)
    
    Call Writer.WriteString8(details)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, _
       ByVal OfferSlot As Byte, _
       ByVal ObjIndex As Integer, _
       ByVal Amount As Long)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
    '25/11/2009: ZaMa - Now sends the specific offer slot to be modified.
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de solo Def
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ChangeUserTradeSlot)
    
    Call Writer.WriteInt8(OfferSlot)
    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteInt32(Amount)
    
    If ObjIndex > 0 Then
        Call Writer.WriteInt32(ObjData(ObjIndex).GrhIndex)
        Call Writer.WriteInt8(ObjData(ObjIndex).OBJType)
        Call Writer.WriteInt16(ObjData(ObjIndex).MaxHIT)
        Call Writer.WriteInt16(ObjData(ObjIndex).MinHIT)
        Call Writer.WriteInt16(ObjData(ObjIndex).MaxDef)
        Call Writer.WriteInt16(ObjData(ObjIndex).MinDef)
        Call Writer.WriteInt32(SalePrice(ObjIndex))
        Call Writer.WriteString8(ObjData(ObjIndex).Name)
        Call Writer.WriteBool(ItemIncompatibleConUser(UserIndex, ObjIndex))
        
    Else        ' Borra el item
        Call Writer.WriteInt32(0)
        Call Writer.WriteInt8(0)
        Call Writer.WriteInt16(0)
        Call Writer.WriteInt16(0)
        Call Writer.WriteInt16(0)
        Call Writer.WriteInt16(0)
        Call Writer.WriteInt32(0)
        Call Writer.WriteString8("")
        Call Writer.WriteBool(False)
        
    End If
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "SendNight" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendNight(ByVal UserIndex As Integer, ByVal night As Boolean)
    
    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Writes the "SendNight" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.SendNight)
    Call Writer.WriteBool(night)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    Dim Tmp         As String
    
    Call Writer.WriteInt8(ServerPacketID.SpawnList)
    
    For i = LBound(npcNames()) To UBound(npcNames())
        Tmp = Tmp & npcNames(i) & SEPARATOR
    Next i
    
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    Dim Tmp         As String
    
    Call Writer.WriteInt8(ServerPacketID.ShowSOSForm)
    
    For i = 1 To Ayuda.Longitud
        Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
    Next i
    
    If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ShowDenounces" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowDenounces(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    'Writes the "ShowDenounces" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim DenounceIndex As Long
    
    Dim DenounceList  As String
    
    Call Writer.WriteInt8(ServerPacketID.ShowDenounces)
    
    For DenounceIndex = 1 To Denuncias.Longitud
        DenounceList = DenounceList & Denuncias.VerElemento(DenounceIndex, False) & SEPARATOR
    Next DenounceIndex
    
    If LenB(DenounceList) <> 0 Then DenounceList = Left$(DenounceList, Len(DenounceList) - 1)
    
    Call Writer.WriteString8(DenounceList)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowPartyForm(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "ShowPartyForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i                         As Long
    
    Dim Tmp                       As String
    
    Dim PI                        As Integer
    
    Dim members(PARTY_MAXMEMBERS) As Integer
    
    Call Writer.WriteInt8(ServerPacketID.ShowPartyForm)
    
    PI = UserList(UserIndex).PartyIndex
    Call Writer.WriteInt8(CByte(Parties(PI).EsPartyLeader(UserIndex)))
    
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(members())
        
        For i = 1 To PARTY_MAXMEMBERS
            
            If members(i) > 0 Then
                Tmp = Tmp & UserList(members(i)).Name & " (" & Fix(Parties(PI).MiExperiencia(members(i))) & ")" & SEPARATOR
                
            End If
            
        Next i
        
    End If
    
    If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    Call Writer.WriteInt32(Parties(PI).ObtenerExperienciaTotal)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, _
       ByVal currentMOTD As String)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ShowMOTDEditionForm)
    
    Call Writer.WriteString8(currentMOTD)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ShowGMPanelForm)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, _
       ByRef userNamesList() As String, _
       ByVal cant As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06 NIGO:
    'Writes the "UserNameList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    Dim Tmp         As String
    
    Call Writer.WriteInt8(ServerPacketID.UserNameList)
    
    ' Prepare user's names list
    For i = 1 To cant
        Tmp = Tmp & userNamesList(i) & SEPARATOR
    Next i
    
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    
    Call Writer.WriteString8(Tmp)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Pong" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.Pong)
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, _
       ByVal invisible As Boolean) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "SetInvisible" message and returns it.
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.SetInvisible)
        
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteBool(invisible)
        
    Exit Function
        
    
End Function

Public Function PrepareMessageCharacterChangeNick(ByVal CharIndex As Integer, _
       ByVal newNick As String) As String
    
    '***************************************************
    'Author: Budi
    'Last Modification: 07/23/09
    'Prepares the "Change Nick" message and returns it.
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.CharacterChangeNick)
        
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteString8(newNick)
        
        Exit Function
    
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal Chat As String, _
       ByVal CharIndex As Integer, _
       ByVal color As Long) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ChatOverHead" message and returns it.
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.ChatOverHead)
        Call Writer.WriteString8(Chat)
        Call Writer.WriteInt16(CharIndex)
        
        ' Write rgb channels and save one byte from long :D
        Call Writer.WriteInt8(color And &HFF)
        Call Writer.WriteInt8((color And &HFF00&) \ &H100&)
        Call Writer.WriteInt8((color And &HFF0000) \ &H10000)
        
    Exit Function
    
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal Chat As String, _
       ByVal FontIndex As FontTypeNames) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ConsoleMsg" message and returns it.
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.ConsoleMsg)
        Call Writer.WriteString8(Chat)
        Call Writer.WriteInt8(FontIndex)
        
    Exit Function
    
End Function
Public Function PrepareRenderConsoleMsg(ByVal Chat As String, _
       ByVal FontIndex As Integer) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ConsoleMsg" message and returns it.
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.RenderMsg)
        Call Writer.WriteString8(Chat)
        Call Writer.WriteInt16(FontIndex)
        
        
    Exit Function
    
End Function
Public Function PrepareCommerceConsoleMsg(ByRef Chat As String, _
       ByVal FontIndex As FontTypeNames) As String
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 03/12/2009
    'Prepares the "CommerceConsoleMsg" message and returns it.
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.CommerceChat)
        Call Writer.WriteString8(Chat)
        Call Writer.WriteInt8(FontIndex)
        
    Exit Function
    
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

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, _
       ByVal FX As Integer, _
       ByVal FXLoops As Integer) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CreateFX" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.CreateFX)
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteInt16(FX)
        Call Writer.WriteInt16(FXLoops)
    Exit Function
    
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Byte, _
       ByVal X As Byte, _
       ByVal Y As Byte) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.PlayWave)
        Call Writer.WriteInt8(wave)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        
    Exit Function
    
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal Chat As String) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "GuildChat" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.GuildChat)
        Call Writer.WriteString8(Chat)
        
    Exit Function
    
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
        Call Writer.WriteInt8(ServerPacketID.ShowMessageBox)
        Call Writer.WriteString8(Chat)
        
    Exit Function
    
End Function

''
' Prepares the "PlayMp3" message and returns it.
'
' @param    mp3 The mp3 to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMp3(ByVal mp3 As Integer, _
       Optional ByVal loops As Integer = -1) As String
    
    '***************************************************
    'Author: Lucas Recoaro (Recox)
    'Last Modification: 05/17/06
    'Prepares the "PlayMp3" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.PlayMp3)
        Call Writer.WriteInt16(mp3)
        Call Writer.WriteInt16(loops)
    Exit Function
    
End Function

''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Integer, _
       Optional ByVal loops As Integer = -1) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "PlayMidi" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.PlayMidi)
        Call Writer.WriteInt16(midi)
        Call Writer.WriteInt16(loops)
        
    Exit Function
    
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "PauseToggle" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.PauseToggle)
    Exit Function
    
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRainToggle() As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "RainToggle" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.RainToggle)
        
    Exit Function
    
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ObjectDelete" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.ObjectDelete)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        
    Exit Function
    
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, _
       ByVal Y As Byte, _
       ByVal Blocked As Boolean) As String
    
    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Prepares the "BlockPosition" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.BlockPosition)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        Call Writer.WriteBool(Blocked)
        
    Exit Function
    
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Long, _
       ByVal X As Byte, _
       ByVal Y As Byte) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'prepares the "ObjectCreate" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.ObjectCreate)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        Call Writer.WriteInt32(GrhIndex)
    Exit Function
    
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterRemove" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.CharacterRemove)
        Call Writer.WriteInt16(CharIndex)
    Exit Function
    
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.RemoveCharDialog)
        Call Writer.WriteInt16(CharIndex)
    Exit Function
    
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
' @param    NickColor Determines if the character is a criminal or not, and if can be atacked by someone
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, _
       ByVal Head As Integer, _
       ByVal heading As eHeading, _
       ByVal CharIndex As Integer, _
       ByVal X As Byte, _
       ByVal Y As Byte, _
       ByVal weapon As Integer, _
       ByVal shield As Integer, _
       ByVal FX As Integer, _
       ByVal FXLoops As Integer, _
       ByVal helmet As Integer, _
       ByVal Name As String, _
       ByVal NickColor As Byte, _
       ByVal Privileges As Byte) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterCreate" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.CharacterCreate)
        
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteInt16(body)
        Call Writer.WriteInt16(Head)
        Call Writer.WriteInt8(heading)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        Call Writer.WriteInt16(weapon)
        Call Writer.WriteInt16(shield)
        Call Writer.WriteInt16(helmet)
        Call Writer.WriteInt16(FX)
        Call Writer.WriteInt16(FXLoops)
        Call Writer.WriteString8(Name)
        Call Writer.WriteInt8(NickColor)
        Call Writer.WriteInt8(Privileges)
        
    Exit Function
        
    
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

Public Function PrepareMessageCharacterChange(ByVal body As Integer, _
       ByVal Head As Integer, _
       ByVal heading As eHeading, _
       ByVal CharIndex As Integer, _
       ByVal weapon As Integer, _
       ByVal shield As Integer, _
       ByVal FX As Integer, _
       ByVal FXLoops As Integer, _
       ByVal helmet As Integer) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterChange" message and returns it
    '***************************************************

        Call Writer.WriteInt8(ServerPacketID.CharacterChange)
        
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteInt16(body)
        Call Writer.WriteInt16(Head)
        Call Writer.WriteInt16(weapon)
        Call Writer.WriteInt16(shield)
        Call Writer.WriteInt16(helmet)
        Call Writer.WriteInt16(FX)
        Call Writer.WriteInt16(FXLoops)
        
    Exit Function
    
End Function
Public Function PrepareMessageHeadingChange(ByVal heading As eHeading, _
       ByVal CharIndex As Integer)
    
    '***************************************************
    'Author: FrankoH298
    'Last Modification: 10/09/19
    'Prepares the "HeadingChange" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.HeadingChange)
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteInt8(heading)
        
    Exit Function
    
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, _
       ByVal X As Byte, _
       ByVal Y As Byte) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterMove" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.CharacterMove)
        Call Writer.WriteInt16(CharIndex)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(Y)
        
    Exit Function
    
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Prepares the "ForceCharMove" message and returns it
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.ForceCharMove)
        Call Writer.WriteInt8(Direccion)
    Exit Function
    
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, _
       ByVal NickColor As Byte, _
       ByRef Tag As String) As String
    
    '***************************************************
    'Author: Alejandro Salvo (Salvito)
    'Last Modification: 04/07/07
    'Last Modified By: Juan Martin Sotuyo Dodero (Maraxus)
    'Prepares the "UpdateTagAndStatus" message and returns it
    '15/01/2010: ZaMa - Now sends the nick color instead of the status.
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.UpdateTagAndStatus)
        
        Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
        Call Writer.WriteInt8(NickColor)
        Call Writer.WriteString8(Tag)
    Exit Function
    
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal Message As String) As String
    
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ErrorMsg" message and returns it
    '***************************************************

        Call Writer.WriteInt8(ServerPacketID.errorMsg)
        Call Writer.WriteString8(Message)
        
    Exit Function
    
End Function

''
' Writes the "StopWorking" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.

Public Sub WriteStopWorking(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 21/02/2010
    '
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.StopWorking)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "CancelOfferItem" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Slot      The slot to cancel.

Public Sub WriteCancelOfferItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
    
    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 05/03/2010
    '
    '***************************************************
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.CancelOfferItem)
    Call Writer.WriteInt8(Slot)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Handles the "SetDialog" message.
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSetDialog(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Amraphen
    'Last Modification: 18/11/2010
    '20/11/2010: ZaMa - Arreglo privilegios.
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim NewDialog As String
        
        NewDialog = Reader.ReadString8()
        
        If .flags.TargetNPC > 0 Then
            
            ' Dsgm/Dsrm/Rm
            If Not ((.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster)) Then
                'Replace the NPC's dialog.
                Npclist(.flags.TargetNPC).Desc = NewDialog
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "Impersonate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImpersonate(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    '
    '***************************************************
    With UserList(UserIndex)
        
        ' Dsgm/Dsrm/Rm
        If (.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim NpcIndex As Integer
        
        NpcIndex = .flags.TargetNPC
        
        If NpcIndex = 0 Then Exit Sub
        
        ' Copy head, body and desc
        Call ImitateNpc(UserIndex, NpcIndex)
        
        ' Teleports user to npc's coords
        Call WarpUserChar(UserIndex, Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, False, True)
        
        ' Log gm
        Call LogGM(.Name, "/IMPERSONAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        
        ' Remove npc
        Call QuitarNPC(NpcIndex)
        
    End With
    
End Sub

''
' Handles the "Imitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImitate(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    '
    '***************************************************
    With UserList(UserIndex)
        
        ' Dsgm/Dsrm/Rm/ConseRm
        If (.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) And (.flags.Privilegios And (PlayerType.consejero Or PlayerType.RoleMaster)) <> (PlayerType.consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim NpcIndex As Integer
        
        NpcIndex = .flags.TargetNPC
        
        If NpcIndex = 0 Then Exit Sub
        
        ' Copy head, body and desc
        Call ImitateNpc(UserIndex, NpcIndex)
        Call LogGM(.Name, "/MIMETIZAR con " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        
    End With
    
End Sub

''
' Handles the "RecordAdd" message.
'
' @param UserIndex The index of the user sending the message

Public Sub HandleRecordAdd(ByVal UserIndex As Integer)
    
    '**************************************************************
    'Author: Amraphen
    'Last Modify Date: 29/11/2010
    '
    '**************************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        Dim Reason   As String
        
        UserName = Reader.ReadString8()
        Reason = Reader.ReadString8()
        
        If Not (.flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.RoleMaster)) Then
            
            'Verificamos que exista el personaje
            If Not PersonajeExiste(UserName) Then
                Call WriteShowMessageBox(UserIndex, "El personaje no existe")
            Else
                'Agregamos el seguimiento
                Call AddRecord(UserIndex, UserName, Reason)
                
                'Enviamos la nueva lista de personajes
                Call WriteRecordList(UserIndex)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "RecordAddObs" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordAddObs(ByVal UserIndex As Integer)
    
    '**************************************************************
    'Author: Amraphen
    'Last Modify Date: 29/11/2010
    '
    '**************************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim RecordIndex As Byte
        
        Dim Obs         As String
        
        RecordIndex = Reader.ReadInt8
        Obs = Reader.ReadString8()
        
        If Not (.flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.RoleMaster)) Then
            'Agregamos la observacion
            Call AddObs(UserIndex, RecordIndex, Obs)
            
            'Actualizamos la informacion
            Call WriteRecordDetails(UserIndex, RecordIndex)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "RecordRemove" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordRemove(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    '
    '***************************************************
    Dim RecordIndex As Integer
    
    With UserList(UserIndex)
        
        RecordIndex = Reader.ReadInt8
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        'Solo dioses pueden remover los seguimientos, los otros reciben una advertencia:
        If (.flags.Privilegios And PlayerType.Dios) Then
            Call RemoveRecord(RecordIndex)
            Call WriteShowMessageBox(UserIndex, "Se ha eliminado el seguimiento.")
            Call WriteRecordList(UserIndex)
        Else
            Call WriteShowMessageBox(UserIndex, "Solo los dioses pueden eliminar seguimientos.")
            
        End If
        
    End With
    
End Sub

''
' Handles the "RecordListRequest" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordListRequest(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    '
    '***************************************************
    With UserList(UserIndex)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call WriteRecordList(UserIndex)
        
    End With
    
End Sub

''
' Writes the "RecordDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordDetails(ByVal UserIndex As Integer, ByVal RecordIndex As Integer)
    
    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    'Writes the "RecordDetails" message to the given user's outgoing data buffer
    '***************************************************
    Dim i           As Long
    
    Dim tIndex      As Integer
    
    Dim tmpStr      As String
    
    Dim TempDate    As Date
    
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.RecordDetails)
    
    'Creador y motivo
    Call Writer.WriteString8(Records(RecordIndex).Creador)
    Call Writer.WriteString8(Records(RecordIndex).Motivo)
    
    tIndex = NameIndex(Records(RecordIndex).Usuario)
    
    'Status del pj (online?)
    Call Writer.WriteBool(tIndex > 0)
    
    'Escribo la IP segUn el estado del personaje
    If tIndex > 0 Then
        'La IP Actual
        tmpStr = UserList(tIndex).IP
    Else        'String nulo
        tmpStr = vbNullString
        
    End If
    
    Call Writer.WriteString8(tmpStr)
    
    'Escribo tiempo online segUn el estado del personaje
    If tIndex > 0 Then
        'Tiempo logueado.
        TempDate = Now - UserList(tIndex).LogOnTime
        tmpStr = Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate)
    Else
        'Envio string nulo.
        tmpStr = vbNullString
        
    End If
    
    Call Writer.WriteString8(tmpStr)
    
    'Escribo observaciones:
    tmpStr = vbNullString
    
    If Records(RecordIndex).NumObs Then
        
        For i = 1 To Records(RecordIndex).NumObs
            tmpStr = tmpStr & Records(RecordIndex).Obs(i).Creador & "> " & Records(RecordIndex).Obs(i).Detalles & vbCrLf
        Next i
        
        tmpStr = Left$(tmpStr, Len(tmpStr) - 1)
        
    End If
    
    Call Writer.WriteString8(tmpStr)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Writes the "RecordList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordList(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    'Writes the "RecordList" message to the given user's outgoing data buffer
    '***************************************************
    Dim i           As Long
    
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.RecordList)
    
    Call Writer.WriteInt8(NumRecords)
    
    For i = 1 To NumRecords
        Call Writer.WriteString8(Records(i).Usuario)
    Next i
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

''
' Handles the "RecordDetailsRequest" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordDetailsRequest(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Amraphen
    'Last Modification: 07/04/2011
    'Handles the "RecordListRequest" message
    '***************************************************
    Dim RecordIndex As Byte
    
    With UserList(UserIndex)
        
        RecordIndex = Reader.ReadInt8
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call WriteRecordDetails(UserIndex, RecordIndex)
        
    End With
    
End Sub

Public Sub HandleMoveItem(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Ignacio Mariano Tirabasso (Budi)
    'Last Modification: 01/01/2011
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        Dim originalSlot As Byte
        
        Dim newSlot      As Byte
        
        originalSlot = Reader.ReadInt8
        newSlot = Reader.ReadInt8
        
        Call InvUsuario.moveItem(UserIndex, originalSlot, newSlot)
        
    End With
    
End Sub

''
' Handles the "LoginExistingAccount" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingAccount(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 12/10/2018
    '
    '***************************************************
    
    On Error GoTo errHandler
    
    Dim UserName    As String
    Dim Password    As String
    Dim version     As String
    
    UserName = Reader.ReadString8()
    Password = Reader.ReadString8()
    
    'Convert version number to string
    version = CStr(Reader.ReadInt8) & "." & CStr(Reader.ReadInt8) & "." & CStr(Reader.ReadInt8)
    
    'If we got here then packet is complete
    
    If Not CuentaExiste(UserName) Then
        Call WriteErrorMsg(UserIndex, "La cuenta no existe.")
        Call CloseSocket(UserIndex)
        Exit Sub
        
    End If
    
    If Not VersionOK(version) Then
        Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la ultima version es la " & ULTIMAVERSION & ". Tu Version " & version & ". La misma se encuentra disponible en www.argentumonline.org")
    Else
        Call ConnectAccount(UserIndex, UserName, Password)
        
    End If
    
    Exit Sub
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "LoginNewAccount" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginNewAccount(ByVal UserIndex As Integer)
    
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/05/2019
    '
    'CHOTS: Fix a bug reported by @juanmz
    '***************************************************
    
    On Error GoTo errHandler
    
    Dim UserName    As String
    Dim Password    As String
    Dim version     As String
    
    UserName = Reader.ReadString8()
    Password = Reader.ReadString8()
    
    If CuentaExiste(UserName) Then
        Call WriteErrorMsg(UserIndex, "La cuenta ya existe.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    'Convert version number to string
    version = CStr(Reader.ReadInt8) & "." & CStr(Reader.ReadInt8) & "." & CStr(Reader.ReadInt8)
    
    If Not VersionOK(version) Then
        Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la ultima version es la " & ULTIMAVERSION & ". Tu Version " & version & ". La misma se encuentra disponible en www.argentumonline.org")
    Else
        Call CreateNewAccount(UserIndex, UserName, Password)
    End If
    
    'If we got here then packet is complete
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

Public Sub WriteUserAccountLogged(ByVal UserIndex As Integer, _
       ByVal UserName As String, _
       ByVal AccountHash As String, _
       ByVal NumberOfCharacters As Byte, _
       ByRef Characters() As AccountUser)
    
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 12/10/2018
    'Writes the "AccountLogged" message to the given user with the data of the account he just logged in
    '***************************************************
    On Error GoTo errHandler
    
    Dim i           As Long
    
    Call Writer.WriteInt8(ServerPacketID.AccountLogged)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(AccountHash)
    Call Writer.WriteInt8(NumberOfCharacters)
    
    If NumberOfCharacters > 0 Then
        
        For i = 1 To NumberOfCharacters
            Call Writer.WriteString8(Characters(i).Name)
            Call Writer.WriteInt16(Characters(i).body)
            Call Writer.WriteInt16(Characters(i).Head)
            Call Writer.WriteInt16(Characters(i).weapon)
            Call Writer.WriteInt16(Characters(i).shield)
            Call Writer.WriteInt16(Characters(i).helmet)
            Call Writer.WriteInt8(Characters(i).Class)
            Call Writer.WriteInt8(Characters(i).race)
            Call Writer.WriteInt16(Characters(i).Map)
            Call Writer.WriteInt8(Characters(i).level)
            Call Writer.WriteInt32(Characters(i).Gold)
            Call Writer.WriteBool(Characters(i).criminal)
            Call Writer.WriteBool(Characters(i).dead)
            Call Writer.WriteBool(Characters(i).gameMaster)
        Next i
        
    End If
    
    Call SaveLastIpsAccountCharfile(UserName, UserList(UserIndex).IP)
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

Public Function PrepareMessagePalabrasMagicas(ByVal SpellIndex As Byte, _
       ByVal CharIndex As Integer) As String
    
    '***************************************************
    '@Shak: Creada el dia 27-08-2016
    'Utilizamos esto para enviar las palabras magicas
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.PalabrasMagicas)
        Call Writer.WriteInt8(SpellIndex)
        Call Writer.WriteInt16(CharIndex)
        
    Exit Function
    
End Function

Public Function PrepareMessageCharacterAttackAnim(ByVal CharIndex As Integer) As String
    
    '***************************************************
    'Author: Cucsijuan
    'Last Modification: 2/9/2018
    'Prepares the Attack animation message.
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.PlayAttackAnim)
        Call Writer.WriteInt16(CharIndex)
        
    Exit Function
    
End Function

Public Function PrepareMessageFXtoMap(ByVal FxIndex As Integer, _
       ByVal loops As Byte, _
       ByVal X As Integer, _
       ByVal Y As Integer) As String
    
        Call Writer.WriteInt8(ServerPacketID.FXtoMap)
        Call Writer.WriteInt8(loops)
        Call Writer.WriteInt16(X)
        Call Writer.WriteInt16(Y)
        Call Writer.WriteInt16(FxIndex)
        
    Exit Function
    
End Function

Public Function WriteSearchList(ByVal UserIndex As Integer, _
       ByVal Num As Integer, _
       ByVal Datos As String, _
       ByVal obj As Boolean) As String
    
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.SearchList)
    Call Writer.WriteInt16(Num)
    Call Writer.WriteBool(obj)
    Call Writer.WriteString8(Datos)
    
errHandler:
    
End Function

Public Sub HandleSearchNpc(ByVal UserIndex As Integer)
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim i       As Long
        
        Dim n       As Integer
        
        Dim Name    As String
        
        Dim UserNpc As String
        
        Dim tStr    As String
        
        UserNpc = Reader.ReadString8()
        
        tStr = Tilde(UserNpc)
        
        For i = 1 To val(LeerNPCs.GetValue("INIT", "NumNPCs"))
            Name = LeerNPCs.GetValue("NPC" & i, "Name")
            
            If InStr(1, Tilde(Name), tStr) Then
                Call WriteSearchList(UserIndex, i, CStr(i & " - " & Name), False)
                n = n + 1
                
            End If
            
        Next i
        
        If n = 0 Then
            Call WriteSearchList(UserIndex, 0, "No hubo resultados de la busqueda.", False)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

Private Sub HandleSearchObj(ByVal UserIndex As Integer)
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserObj As String
        
        Dim tUser   As Integer
        
        Dim n       As Integer
        
        Dim i       As Long
        
        Dim tStr    As String
        
        UserObj = Reader.ReadString8()
        
        tStr = Tilde(UserObj)
        
        For i = 1 To UBound(ObjData)
            
            If InStr(1, Tilde(ObjData(i).Name), tStr) Then
                Call WriteSearchList(UserIndex, i, CStr(i & " - " & ObjData(i).Name), True)
                n = n + 1
                
            End If
            
        Next
        
        If n = 0 Then
            Call WriteSearchList(UserIndex, 0, "No hubo resultados de la busqueda.", False)
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

Private Sub HandleEnviaCvc(ByVal UserIndex As Integer)
    
    'Dim targetIndex As Integer
    
    With UserList(UserIndex)
        
        If .flags.TargetUser = 0 Then Exit Sub        'gdk: adonde mierda clickeas manko
        Call Mod_ClanvsClan.Enviar(UserIndex, .flags.TargetUser)
        
    End With
    
End Sub

Private Sub HandleAceptarCvc(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        
        If .flags.TargetUser = 0 Then Exit Sub
        Call Mod_ClanvsClan.Aceptar(UserIndex, .flags.TargetUser)
        
    End With
    
End Sub

Private Sub HandleIrCvc(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        
        Call Mod_ClanvsClan.ConectarCVC(UserIndex, True)        'gdk: si le pones false bugeas toditus.
        
    End With
    
End Sub

Public Sub HandleDragAndDropHechizos(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        
        Dim AnteriorPosicion As Integer: AnteriorPosicion = Reader.ReadInt16
        Dim NuevaPosicion As Integer: NuevaPosicion = Reader.ReadInt16
        
        Dim Hechizo As Integer: Hechizo = .Stats.UserHechizos(NuevaPosicion)
        
        .Stats.UserHechizos(NuevaPosicion) = .Stats.UserHechizos(AnteriorPosicion)
        .Stats.UserHechizos(AnteriorPosicion) = Hechizo
        
    End With
    
End Sub

Public Sub WriteQuestDetails(ByVal UserIndex As Integer, _
       ByVal QuestIndex As Integer, _
       Optional QuestSlot As Byte = 0)
    
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Envï¿½a el paquete QuestDetails y la informaciï¿½n correspondiente.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i           As Integer
    
    On Error GoTo errHandler
    
    'ID del paquete
    Call Writer.WriteInt8(ServerPacketID.QuestDetails)
    
    'Se usa la variable QuestSlot para saber si enviamos la info de una quest ya empezada o la info de una quest que no se aceptï¿½ todavï¿½a (1 para el primer caso y 0 para el segundo)
    Call Writer.WriteInt8(IIf(QuestSlot, 1, 0))
    
    'Enviamos nombre, descripciï¿½n y nivel requerido de la quest
    Call Writer.WriteString8(QuestList(QuestIndex).Nombre)
    Call Writer.WriteString8(QuestList(QuestIndex).Desc)
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredLevel)
    
    'Enviamos la cantidad de npcs requeridos
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredNPCs)
    
    If QuestList(QuestIndex).RequiredNPCs Then
        
        'Si hay npcs entonces enviamos la lista
        For i = 1 To QuestList(QuestIndex).RequiredNPCs
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).Amount)
            Call Writer.WriteString8(GetVar(DatPath & "NPCs.dat", "NPC" & QuestList(QuestIndex).RequiredNPC(i).NpcIndex, "Name"))
            
            'Si es una quest ya empezada, entonces mandamos los NPCs que matï¿½.
            If QuestSlot Then
                Call Writer.WriteInt16(UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i))
                
            End If
            
        Next i
        
    End If
    
    'Enviamos la cantidad de objs requeridos
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredOBJs)
    
    If QuestList(QuestIndex).RequiredOBJs Then
        
        'Si hay objs entonces enviamos la lista
        For i = 1 To QuestList(QuestIndex).RequiredOBJs
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).Amount)
            Call Writer.WriteString8(ObjData(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex).Name)
        Next i
        
    End If
    
    'Enviamos la recompensa de oro y experiencia.
    Call Writer.WriteInt32(QuestList(QuestIndex).RewardGLD)
    Call Writer.WriteInt32(QuestList(QuestIndex).RewardEXP)
    
    'Enviamos la cantidad de objs de recompensa
    Call Writer.WriteInt8(QuestList(QuestIndex).RewardOBJs)
    
    If QuestList(QuestIndex).RewardOBJs Then
        
        'si hay objs entonces enviamos la lista
        For i = 1 To QuestList(QuestIndex).RewardOBJs
            Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).Amount)
            Call Writer.WriteString8(ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).Name)
        Next i
        
    End If
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

Public Sub WriteQuestListSend(ByVal UserIndex As Integer)
    
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Envï¿½a el paquete QuestList y la informaciï¿½n correspondiente.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i           As Integer
    
    Dim tmpStr      As String
    
    Dim tmpByte     As Byte
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        Writer.WriteInt8 ServerPacketID.QuestListSend
        
        For i = 1 To MAXUSERQUESTS
            
            If .QuestStats.Quests(i).QuestIndex Then
                tmpByte = tmpByte + 1
                tmpStr = tmpStr & QuestList(.QuestStats.Quests(i).QuestIndex).Nombre & "-"
                
            End If
            
        Next i
        
        'Escribimos la cantidad de quests
        Call Writer.WriteInt8(tmpByte)
        
        'Escribimos la lista de quests (sacamos el ï¿½ltimo caracter)
        If tmpByte Then
            Call Writer.WriteString8(Left$(tmpStr, Len(tmpStr) - 1))
            
        End If
        
    End With
    
    Exit Sub
    
errHandler:
    
    Call Writer.Clear
    
End Sub

Public Function PrepareMessageCreateDamage(ByVal X As Byte, ByVal Y As Byte, ByVal DamageValue As Long, ByVal DamageType As Byte)
    
    ' @ Envia el paquete para crear dano (Y)
    
        Writer.WriteInt8 ServerPacketID.CreateDamage
        Writer.WriteInt8 X
        Writer.WriteInt8 Y
        Writer.WriteInt32 DamageValue
        Writer.WriteInt8 DamageType
        
    Exit Function
End Function

Public Sub HandleCambiarContrasena(ByVal UserIndex As Integer)
    
    'Verifico si llegan todos los datos
    
    On Error GoTo errHandler
    
    Dim Correo      As String
    Dim NuevaContrasena As String
    
    With UserList(UserIndex)
        
        'Leemos los datos de la cuenta a modificar.
        Correo = Reader.ReadString8()
        NuevaContrasena = Reader.ReadString8()
        
        If ConexionAPI Then
            
            'Correo = UserName es lo mismo para aca el Jopi le puso correo :)
            If Not CuentaExiste(Correo) Then
                Call WriteErrorMsg(UserIndex, "La cuenta no existe.")
                Call CloseSocket(UserIndex)
                Exit Sub
                
            End If
            
            Call ApiEndpointSendResetPasswordAccountEmail(Correo, NuevaContrasena)
            
            Call WriteErrorMsg(UserIndex, "Se ha enviado un correo electronico a: " & Correo & " donde debera confirmar el cambio de la password de su cuenta.")
            
        Else
            
            Call WriteErrorMsg(UserIndex, "Esta funcion se encuentra deshabilitada actualmente, si sos el administrador del servidor necesitas habilitar la API hecha en Node.js (https://github.com/ao-libre/ao-api-server).")
            
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        'Por ultimo limpia el buffer nunca poner exit sub antes de limpiar el buffer porque explota
        Call CloseSocket(UserIndex)
        
    End With
    
errHandler:
    
    Dim Error       As Long: Error = Err.Number
    
    Call CloseSocket(UserIndex)
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

Public Sub WriteUserInEvent(ByVal UserIndex As Integer)
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.UserInEvent)
    Exit Sub
    
errHandler:
End Sub

Private Sub HandleFightSend(ByVal UserIndex As Integer)
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim ListUsers As String
        Dim GldRequired As Long
        Dim Users() As String
        
        ListUsers = Reader.ReadString8() & "-" & .Name
        GldRequired = Reader.ReadInt32
        
        If Len(ListUsers) >= 1 Then
            Users = Split(ListUsers, "-")
            
            Call Retos.SendFight(UserIndex, GldRequired, Users)
        End If
        
    End With
    
errHandler:
    Dim Error       As Long
    Error = Err.Number
    On Error GoTo 0
    
    If Error <> 0 Then _
       Err.Raise Error
End Sub

Private Sub HandleFightAccept(ByVal UserIndex As Integer)
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim UserName As String
        
        UserName = Reader.ReadString8()
        
        If Len(UserName) >= 1 Then
            Call Retos.AcceptFight(UserIndex, UserName)
        End If
        
    End With
    
errHandler:
    Dim Error       As Long
    Error = Err.Number
    On Error GoTo 0
    
    If Error <> 0 Then _
       Err.Raise Error
End Sub

Private Sub HandleCloseGuild(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        
        Dim i       As Long
        Dim PreviousGuildIndex  As Integer
        
        If Not .GuildIndex >= 1 Then
            Call WriteConsoleMsg(UserIndex, "No perteneces a ningun clan.", FONTTYPE_GUILD)
            Exit Sub
            
        End If
        
        If guilds(.GuildIndex).Fundador <> .Name Then
            Call WriteConsoleMsg(UserIndex, "No eres lider del clan.", FONTTYPE_GUILD)
            Exit Sub
            
        End If
        
        'Ya con cambiarle el nombre a "CLAN CERRADO" ya se omite de la lista de clanes enviadas al cliente.
        'Tambien cambiamos "Founder" y "Leader" a "NADIE" sino no te deja fundar otro clan.
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & .GuildIndex, "GuildName", "CLAN CERRADO")
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & .GuildIndex, "Founder", "NADIE")
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & .GuildIndex, "Leader", "NADIE")
        
        PreviousGuildIndex = .GuildIndex
        
        'Obtenemos la lista de miembros del clan.
        Dim GuildMembers() As String
        GuildMembers = guilds(PreviousGuildIndex).GetMemberList()
        
        For i = 0 To UBound(GuildMembers)
            Call SaveUserGuildIndex(GuildMembers(i), 0)
            Call SaveUserGuildAspirant(GuildMembers(i), 0)
        Next i
        
        'La borramos junto con la lista de solicitudes.
        Call Kill(App.Path & "\Guilds\" & guilds(PreviousGuildIndex).GuildName & "-members.mem")
        Call Kill(App.Path & "\Guilds\" & guilds(PreviousGuildIndex).GuildName & "-solicitudes.sol")
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El Clan " & guilds(.GuildIndex).GuildName & " ha cerrado sus puertas.", FontTypeNames.FONTTYPE_GUILD))
        
    End With
    
    ' Actualizamos la base de datos de clanes.
    Call modGuilds.LoadGuildsDB
    
    Exit Sub
    
End Sub

''
' Handles the "Discord" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDiscord(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Lucas Daniel Recoaro (Recox)
    'Last Modification: 14/07/19 (Recox)
    'Manda un mensaje al server para que el mismo lo envie al bot del discord (Recox)
    '***************************************************
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Dim Chat    As String
        Chat = Reader.ReadString8()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            'Aqui solo vamos a hacer un request a los endpoints de la aplicacion en Node.js
            'el repositorio para hacer funcionar esto, es este: https://github.com/ao-libre/ao-api-server
            'Si no tienen interes en usarlo pueden desactivarlo en el Server.ini
            If ConexionAPI Then
                
                Call ApiEndpointSendCustomCharacterMessageDiscord(Chat, .Name, .Desc)
                Call WriteConsoleMsg(UserIndex, "Link Discord: https://discord.gg/xbAuHcf - El bot de Discord recibio y envio lo siguiente: " & Chat, FontTypeNames.FONTTYPE_INFOBOLD)
                
            Else
                Call WriteConsoleMsg(UserIndex, "(api - node.js)  El modulo para usar esta funcion no esta instalado en este servidor. http://www.github.com/ao-libre/ao-api-server para mas informacion / more info.", FontTypeNames.FONTTYPE_INFOBOLD)
                
            End If
            
        End If
        
    End With
    
errHandler:
    
    Dim Error       As Long
    
    Error = Err.Number
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

Public Sub HandleLimpiarMundo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Jopi
    'Last Modification: 06/07/2020
    'Fuerza una limpieza del mundo.
    '***************************************************
    
    'Me fijo si es GM
    If Not EsGm(UserIndex) Then Exit Sub
    
    Call LogGM(UserList(UserIndex).Name, "forzo la limpieza del mundo.")
    
    tickLimpieza = 16
    
End Sub

''
' Writes the "EquitandoToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquitandoToggle(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 23/08/11
    'Writes the "EquitandoToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    With UserList(UserIndex)
        Call Writer.WriteInt8(ServerPacketID.EquitandoToggle)
        
        Call Writer.WriteInt32(.Counters.MonturaCounter)
        
    End With
    
    Exit Sub
    
errHandler:
    
End Sub

Private Sub HandleObtenerDatosServer(ByVal UserIndex As Integer)
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        
        Call WriteEnviarDatosServer(UserIndex)
        
    End With
    
errHandler:
    
    Dim Error       As Long: Error = Err.Number
    
    Call CloseSocket(UserIndex)
    
    On Error GoTo 0
    
    If Error <> 0 Then Err.Raise Error
    
End Sub

''
' Handles the "CraftsmanCreate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCraftsmanCreate(ByVal UserIndex As Integer)
    '***************************************************
    'Author: WyroX
    'Last Modification: 27/01/2020
    '***************************************************
    
    Dim Item        As Integer
    
    Item = Reader.ReadInt16()
    If Item < LBound(ObjArtesano) Or Item > UBound(ObjArtesano) Then Exit Sub
    
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
        Exit Sub
    End If
    
    'El target es un NPC valido
    If UserList(UserIndex).flags.TargetNPC = 0 Then Exit Sub
    
    'Es el artesano
    If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Artesano Then Exit Sub
    
    'Esta cerca
    If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del artesano.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    'User retira el item del slot
    Call ArtesanoConstruirItem(UserIndex, Item)
    
End Sub

Private Sub WriteEnviarDatosServer(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Recox
    'Last Modification: 19/01/20
    'Writes the "WriteEnviarDatosServer" message to the given user's outgoing data buffer
    'Mandamos informacion del server al cliente (Recox)
    '***************************************************
    Dim MundoSeleccionadoWithoutPath As String
    MundoSeleccionadoWithoutPath = Replace(MundoSeleccionado, "\Mundos\", "")
    MundoSeleccionadoWithoutPath = Replace(MundoSeleccionadoWithoutPath, "\", "")
    
    With UserList(UserIndex)
        
        Call Writer.WriteInt8(ServerPacketID.EnviarDatosServer)
        Call Writer.WriteString8(MundoSeleccionadoWithoutPath)
        Call Writer.WriteString8(NombreServidor)
        Call Writer.WriteString8(DescripcionServidor)
        
        ' Mandamos tambien el nivel maximo del server, ya que esto puede variar de servidor en servidor (Recox)
        Call Writer.WriteInt16(STAT_MAXELV)
        
        ' Mandamos maxima cantidad de usuarios simultaneos soportada por el servidor.
        Call Writer.WriteInt16(MaxUsers)
        
        ' Mandamos cantidad usuarios online. Ponemos -1 sino nos cuenta a nosotros mismos y no es la idea
        Call Writer.WriteInt16(LastUser - 1)
        
        ' Mandamos multiplicadores de oro, exp y trabajo
        Call Writer.WriteInt16(ExpMultiplier)
        Call Writer.WriteInt16(OroMultiplier)
        Call Writer.WriteInt16(OficioMultiplier)
        
        'If we got here then packet is complete, copy data back to original queue
        'Por ultimo limpia el buffer nunca poner exit sub antes de limpiar el buffer porque explota
        Call CloseSocket(UserIndex)
        
    End With
    
End Sub

Public Sub WriteCargarListaDeAmigos(ByVal UserIndex As Integer, ByVal Slot As Byte)
    
    On Error GoTo errHandler
    
    Dim i           As Integer
    
    Call Writer.WriteInt8(ServerPacketID.EnviarListDeAmigos)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteString8(UserList(UserIndex).Amigos(Slot).Nombre)
    
    Exit Sub
    
errHandler:
    
End Sub

Public Sub WriteSeeInProcess(ByVal UserIndex As Integer)
    '***************************************************
    'Author:Franco Emmanuel GimÃ©nez (Franeg95)
    'Last Modification: 18/10/10
    '***************************************************
    On Error GoTo errHandler
    Call Writer.WriteInt8(ServerPacketID.SeeInProcess)
    
    Exit Sub
    
errHandler:
End Sub

Private Sub HandleSendProcessList(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Franco Emmanuel Gimenez(Franeg95)
    'Last Modification: 18/10/10
    '***************************************************
    
    On Error GoTo errHandler
    With UserList(UserIndex)
        Dim Captions As String, Process As String
        
        Captions = Reader.ReadString8()
        Process = Reader.ReadString8()
        
        If .flags.GMRequested > 0 Then
            If UserList(.flags.GMRequested).ConnIDValida Then
                Call WriteShowProcess(.flags.GMRequested, Captions, Process)
                .flags.GMRequested = 0
            End If
        End If
        
    End With
    
errHandler:            Dim Error As Long:     Error = Err.Number: On Error GoTo 0:  If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleLookProcess(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Franco Emmanuel Gimenez(Franeg95)
    'Last Modification: Cuicui - 20/07/20
    '***************************************************
    
    On Error GoTo errHandler
    With UserList(UserIndex)
        
        Dim tName   As String
        Dim tIndex  As Integer
        
        tName = Reader.ReadString8()
        
        If EsGm(UserIndex) Then
            tIndex = NameIndex(tName)
            If tIndex > 0 Then
                UserList(tIndex).flags.GMRequested = UserIndex
                Call WriteSeeInProcess(tIndex)
            End If
        End If
        
    End With
    
    Exit Sub
    
errHandler:
    Dim Error       As Long
    Error = Err.Number
    On Error GoTo 0
    If Error <> 0 Then Err.Raise Error
    
    LogError ("Error en HandleLookProcess. Error: " & Err.Number & " - " & Err.description)
End Sub

Public Sub WriteShowProcess(ByVal gmIndex As Integer, ByVal strCaptions As String, ByVal strProcess As String)
    
    On Error GoTo errHandler
    
    Call Writer.WriteInt8(ServerPacketID.ShowProcess)
    Call Writer.WriteString8(strCaptions)
    Call Writer.WriteString8(strProcess)
    
    Exit Sub
errHandler:
    
End Sub

Public Function PrepareMessageProyectil(ByVal UserIndex As Integer, ByVal CharSending As Integer, ByVal CharRecieved As Integer, ByVal GrhIndex As Integer) As String
    '*************************************
    'Autor: Lorwik
    'Last Modification: 12/07/2020
    '*************************************
    
        Writer.WriteInt8 (ServerPacketID.proyectil)
        Writer.WriteInt16 (CharSending)
        Writer.WriteInt16 (CharRecieved)
        Writer.WriteInt16 (GrhIndex)
    Exit Function
    
End Function

Public Function PrepareMessageCharacterIsInChatMode(ByVal CharIndex As Integer) As String
    
    '***************************************************
    'Author: Recox
    'Last Modification: 2/9/2018
    'Prepares the InChatMode animation message.
    '***************************************************
        Call Writer.WriteInt8(ServerPacketID.PlayIsInChatMode)
        Call Writer.WriteInt16(CharIndex)
        
    Exit Function
    
End Function

Private Sub HandleSendIfCharIsInChatMode(ByVal UserIndex As Integer)
    
1           On Error GoTo HandleSendIfCharIsInChatMode_Error
    
2           With UserList(UserIndex)
    
3
8               .Char.Escribiendo = IIf(.Char.Escribiendo = 1, 0, 1)
9               Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageSetTypingFlagToCharIndex(.Char.CharIndex, .Char.Escribiendo))
    
10          End With
    
11          Exit Sub
    
HandleSendIfCharIsInChatMode_Error:
    
12          Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure HandleSendIfCharIsInChatMode of Modulo Protocol " & Erl & ".")
    
End Sub

Private Function PrepareMessageSetTypingFlagToCharIndex(ByVal CharIndex As Integer, _
        ByVal Escribiendo As Byte) As String
    On Error Resume Next
    
    Call Writer.WriteInt8(ServerPacketID.PlayIsInChatMode)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteInt8(Escribiendo)
    
    Exit Function
    
End Function
