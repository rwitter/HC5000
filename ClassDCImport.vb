Option Strict On

Class classDCImport

#Region "Constants"
    Const SUPPORTED_MODELS As String = "{HC5000}"
    Const SUPPORTED_VERSION_MIN As Double = 2.0
#End Region

#Region "Enums"
    Public Enum DoorOption
        LockTime
        OpenTime
        TimeProfile
        TradesMan
    End Enum
    Public Enum DCOption
        ChannelInterLock
        DateValidation
        AntiPassBack
    End Enum
    Private Enum ParserState
        Idle
        Version
        Import
        DoorControllerOptions
        DoorOptions
        EndImport
    End Enum
#End Region

#Region "Structures"
    Private Structure extDoor
        Public LockTime As Integer
        Public OpenTime As Integer
        Public TimeProfile As Integer
        Public TradesMan As Integer
    End Structure

    'External door controller parameters
    Private Structure extDoorController
        Public Model As String
        Public Version As Double
        Public Address As Byte
        Public RecordCount As Integer
        Public NumNoResponses As Integer
        Public Baudrate As Integer                    'BaudRate of Door controller
        Public SupportedModel As Boolean
        Public SupportedVersion As Boolean

        Public BlockID As Integer
        Public DoorControllerID As Integer

        Public InterMessageDelay As Integer             'MS
        Public Timeout As Integer                       'MS

        Public DCOptions As Byte

        Public Door1Options As extDoor
        Public Door2Options As extDoor

    End Structure
#End Region

#Region "Class Variables"
    Private bCloseDown As Boolean = False
    Private CurrentState As ParserState = ParserState.Idle
    'Create a door controller 
    Private DCInfo As extDoorController
    'RS-232 Class
    Private WithEvents Port As RS232
    'RS-232 related events
    Public Event RS232_Data()
    Public Event RS232_NoResponse(ByRef NumNoResponses As Integer)
    Public Event RS232_RetriesExceeded()
    Public Event RS232_TimeOut(ByRef RetryCount As Integer, ByRef LastTxMsg As String)
    'Message parsing
    Public Parser As New ClassHCParse
    Friend LastRxMsg As New classHCMsg
    Private myLastIntRxMsg As New classHCMsg
    'Events
    Public Event VersionRetreived()
    Public Event RecordRetreived(ByRef Msg As classHCMsg, ByRef RecordCount As Integer)
    Public Event NoMoreCardData()
#End Region

#Region "Properties"
    'Door Controller Version
    ReadOnly Property Version() As Double
        Get
            Return DCInfo.Version
        End Get
    End Property
    'Model
    ReadOnly Property Model() As String
        Get
            Return DCInfo.Model
        End Get
    End Property
    'Supported Version
    ReadOnly Property SupportedVersion() As Boolean
        Get
            Return DCInfo.SupportedVersion
        End Get
    End Property
    'Supported Model
    ReadOnly Property SupportedModel() As Boolean
        Get
            Return DCInfo.SupportedModel
        End Get
    End Property
    'Number of records retreived
    ReadOnly Property numRecsGood() As Integer
        Get
            Return DCInfo.RecordCount
        End Get
    End Property
    'Number of Records not retreived
    ReadOnly Property numRecsBad() As Integer
        Get
            Return DCInfo.NumNoResponses
        End Get
    End Property
    'Address of the Door controller
    Property address() As Byte
        Get
            Return DCInfo.Address
        End Get
        Set(ByVal value As Byte)
            DCInfo.Address = value
        End Set
    End Property
    'Baudrate that the Door controller is using
    Property Baudrate() As Integer
        Get
            Return DCInfo.Baudrate
        End Get
        Set(ByVal value As Integer)
            DCInfo.Baudrate = value
        End Set
    End Property
    'BlockID
    Property BlockID() As Integer
        Get
            Return DCInfo.BlockID
        End Get
        Set(ByVal value As Integer)
            DCInfo.BlockID = value
        End Set
    End Property
    'Door Controller ID
    Property DoorControllerID() As Integer
        Get
            Return DCInfo.DoorControllerID
        End Get
        Set(ByVal value As Integer)
            DCInfo.DoorControllerID = value
        End Set
    End Property
    'Delay between sending import messages
    Property InterMessageDelay() As Integer
        Get
            Return DCInfo.InterMessageDelay
        End Get
        Set(ByVal value As Integer)
            DCInfo.InterMessageDelay = value
            Port.InterMessageDelay = value
        End Set
    End Property
    'Timeout for Importing records
    Property ReadTimeOut() As Integer
        Get
            Return DCInfo.Timeout
        End Get
        Set(ByVal value As Integer)
            DCInfo.Timeout = value
            'Set the timeout on the Port
            If Not (Port Is Nothing) Then
                Port.ReadTimeOut = value
            End If
        End Set
    End Property
    'DC Options
    ReadOnly Property DoorControllerOptions(ByVal ReqdOption As DCOption) As Boolean
        Get
            With DCInfo
                Select Case ReqdOption
                    Case DCOption.ChannelInterLock
                        Return CBool(.DCOptions And 1)
                    Case DCOption.AntiPassBack
                        Return CBool(.DCOptions And 2)
                    Case DCOption.DateValidation
                        Return CBool(.DCOptions And 4)
                End Select
            End With
        End Get
    End Property
    'Return Door options for door 1 or 2
    'NB: Tradesman s both for same doors
    ReadOnly Property DoorOptions(ByVal DoorNum As Integer, ByVal ReqdOption As DoorOption) As Byte
        Get
            With DCInfo
                If DoorNum = 1 Then
                    Select Case ReqdOption
                        Case DoorOption.LockTime
                            Return CByte(.Door1Options.LockTime)
                        Case DoorOption.OpenTime
                            Return CByte(.Door1Options.OpenTime)
                        Case DoorOption.TimeProfile
                            Return CByte(.Door1Options.TimeProfile)
                        Case DoorOption.TradesMan
                            Return CByte(.Door1Options.TradesMan)
                    End Select
                End If
                If DoorNum = 2 Then
                    Select Case ReqdOption
                        Case DoorOption.LockTime
                            Return CByte(.Door2Options.LockTime)
                        Case DoorOption.OpenTime
                            Return CByte(.Door2Options.OpenTime)
                        Case DoorOption.TimeProfile
                            Return CByte(.Door2Options.TimeProfile)
                        Case DoorOption.TradesMan
                            Return CByte(.Door2Options.TradesMan)
                    End Select
                End If
            End With
        End Get
    End Property
#End Region

#Region "Main Parser"
    Sub ParseRxMsg()

        Console.WriteLine(CurrentState)

        'Go to idle on import cancellation
        'If bCloseDown And CurrentState = ParserState.Import Then
        ' bCloseDown = Not bCloseDown
        ' CurrentState = ParserState.Idle
        '       End If
        '
        'Main State Machine 
        With LastRxMsg

            Select Case CurrentState
                Case ParserState.Idle
                    CurrentState = ParserState.Version


                Case ParserState.Version
                    If .eFormat = etypeHC_ReplyFormat.HC_DATA_SEND Then
                        Select Case .eType
                            Case etypeHC_Cmd.HC_GET_VERSION
                                InitDCVersionInfo()

                                DCInfo.RecordCount = 0
                                RaiseEvent VersionRetreived()


                        End Select

                        If .eType = etypeHC_Cmd.HC_GET_VERSION Then
                            'Go get door controller options if required
                            If DCInfo.Version >= SUPPORTED_VERSION_MIN Then
                                CurrentState = ParserState.DoorControllerOptions
                                GetDCOptions()
                            End If
                        End If
                    End If

                Case ParserState.DoorControllerOptions
                    Select Case .eFormat
                        Case etypeHC_ReplyFormat.HC_DATA_SEND
                            Select Case .eType
                                Case etypeHC_Cmd.HC_DC_CONFIG
                                    SetDCOptions()
                                    CurrentState = ParserState.DoorOptions
                                    GetDoorOptions()
                            End Select
                    End Select

                Case ParserState.DoorOptions
                    Select Case .eFormat
                        Case etypeHC_ReplyFormat.HC_DATA_SEND
                            Select Case .eType
                                Case etypeHC_Cmd.HC_DOOR_CONFIG
                                    SetDoorOptions()
                                    CurrentState = ParserState.Import
                            End Select
                    End Select

                Case ParserState.Import

                    Select Case .eFormat
                        Case etypeHC_ReplyFormat.HC_DATA_SEND
                            Select Case .eType
                                Case etypeHC_Cmd.HC_DOWNLOAD_USERS
                                    DCInfo.RecordCount += 1
                                    RaiseEvent RecordRetreived(LastRxMsg, DCInfo.RecordCount)
                                    SendAck()
                                Case etypeHC_Cmd.HC_DOOR_CONTROL

                                Case Else
                                    MsgBox("Unaccounted State")
                            End Select

                        Case etypeHC_ReplyFormat.HC_NO_MORE_DATA
                            RaiseEvent NoMoreCardData()
                            'Go back to idle
                            CurrentState = ParserState.Idle
                        Case etypeHC_ReplyFormat.HC_CMD_ACK
                        Case Else
                            MsgBox("Unaccounted State")
                    End Select

                Case Else
                    MsgBox("Unaccounted State")
            End Select
        End With
    End Sub

    Sub ResetParser()
        CurrentState = ParserState.Idle
    End Sub


#End Region

#Region "RS-232"
    Sub DCCom_Init(ByRef COMPort As Byte, ByRef Baudrate As Integer)
        DCCom_Open(COMPort, Baudrate)
        With Port
            .WriteTimeOut = DCInfo.Timeout          'Set the timeouts on the Port
            .ReadTimeOut = DCInfo.Timeout
        End With

    End Sub
    Private Function DCCom_Open(ByRef PortNum As Byte, ByRef Baud As Integer) As Boolean
        'Attempt to open port, based on user options
        '------------------------------------------------
        ' Opens a com port object for data transmission
        '------------------------------------------------
        Try
            If Port Is Nothing Then
                Port = New RS232(PortNum, Baud)   'Set up COM port with selected number
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error: OpenRS232", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If Not (Port Is Nothing) Then Port = Nothing
        End Try
    End Function
    Public ReadOnly Property DCCom_PortIsOpen() As Boolean
        Get
            If Not (Port Is Nothing) Then
                Return Port.IsOpen
            End If
        End Get
    End Property
    Property DCCom_PortExists() As Boolean
        Get
            If Not (Port Is Nothing) Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(ByVal value As Boolean)

        End Set
    End Property
    Sub DCCom_Close()
        Port.Close()
        ResetParser()
    End Sub

#End Region

#Region "RS-232 Events"
    Private Sub objPort_RS232_Data() Handles Port.RS232_Data
        Console.WriteLine("Msg:" & Port.LastRx_String)
        If Parser.ParseRxMessage(Port.LastRx_String) Then                  'Parse the last received message
            Parser.GetLastMessage(LastRxMsg)                            'If is a valid response
            ParseRxMsg()                                                   'Parse the response
            RaiseEvent RS232_Data()
        End If
    End Sub

    Private Sub objPort_RS232_NoResponse() Handles Port.RS232_NoResponse
        'No response -  from a message with no timeout ie; request for bulk upload data
        'Maybe design state machine and send a nack if in bulk upload mode?
        DCInfo.NumNoResponses += 1
        RaiseEvent RS232_NoResponse(DCInfo.NumNoResponses)
        ' Dim LastMsg As String = LastRxMsg.objMessage.ToString
        Select Case LastRxMsg.eFormat
            Case etypeHC_ReplyFormat.HC_CMD_NACK
                'Acknowledge the NACK
                SendAck()
            Case etypeHC_ReplyFormat.HC_DATA_SEND
                SendAck()
        End Select
    End Sub
    Private Sub objPort_RS232_RetriesExceeded() Handles Port.RS232_RetriesExceeded
        RaiseEvent RS232_RetriesExceeded()
    End Sub
    Private Sub objPort_RS232_Timeout() Handles Port.RS232_Timeout_Rx
        If CurrentState = ParserState.Import Then
            SendAck()
        Else
            'Retry last message
            RaiseEvent RS232_TimeOut(Port.RetryCount, Port.LastTX_String)
            SendLastMsg()
        End If
    End Sub
#End Region

#Region "Comms - Message Generation & transmission"
    Sub GetDoorOptions()
        'Set State
        CurrentState = ParserState.DoorOptions
        If ChkPortStatus() Then Port.Transmit(Parser.MakeGetDoorOptions(DCInfo.Address))
    End Sub
    Sub GetDCOptions()
        'Set State
        CurrentState = ParserState.DoorControllerOptions
        If ChkPortStatus() Then Port.Transmit(Parser.MakeGetDCOptions(DCInfo.Address))
    End Sub
    'Retrieve version number
    Sub GetVersion()
        'Initialise State Machine
        CurrentState = ParserState.Version
        If ChkPortStatus() Then Port.Transmit(Parser.MakeGetVersion(DCInfo.Address))
    End Sub
    'Import records from the door controller
    Sub Import()
        If ChkPortStatus() Then
            With DCInfo
                Port.Transmit(Parser.MakeGetCardData(.Address))
            End With
        End If
    End Sub
    'Ack
    Sub SendAck()
        If ChkPortStatus() Then Port.Transmit(Parser.MakeAck(etypeHC_Cmd.HC_DOWNLOAD_USERS), False)
    End Sub
    'NAck
    Sub SendNAck()
        If ChkPortStatus() Then Port.Transmit(Parser.MakeNAck(etypeHC_Cmd.HC_DOWNLOAD_USERS), False)
    End Sub
    'Last Message
    Sub SendLastMsg()
        If ChkPortStatus() Then Port.Transmit(Port.LastTX_StringBuilder, False)
    End Sub
    'Prepares port for transmission, returning status 
    Function ChkPortStatus() As Boolean
        If Not Port Is Nothing Then
            Port.Open()
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region "Response parsing"
    Sub InitDCVersionInfo()
        'Initialises the Door controller versiona nd model and 
        'sets flags to indicate whether the model can support
        'DC import features
        With DCInfo
            .Version = CDbl(LastRxMsg.SubString(13, 4))
            .Model = CStr(LastRxMsg.SubString(3, 6))

            If SUPPORTED_MODELS.Contains(.Model) Then
                .SupportedModel = True
            Else
                .SupportedModel = False
            End If

            If .Version >= SUPPORTED_VERSION_MIN Then
                .SupportedVersion = True
            Else
                .SupportedVersion = False
            End If
        End With
    End Sub
    Sub SetDCOptions()
        Dim Temp As Long
        ParseHex(LastRxMsg.objMessage, 2, 2, Temp)
        DCInfo.DCOptions = CByte(Temp)
    End Sub
    Sub SetDoorOptions()
        Dim Temp As Long
        'Door 1
        'Lock Time
        ParseHex(LastRxMsg.objMessage, 4, 2, Temp)
        DCInfo.Door1Options.LockTime = CByte(Temp)
        'Open Time
        ParseHex(LastRxMsg.objMessage, 6, 2, Temp)
        DCInfo.Door1Options.OpenTime = CByte(Temp)
        'Time Profile
        ParseHex(LastRxMsg.objMessage, 8, 2, Temp)
        DCInfo.Door1Options.TimeProfile = CByte(Temp)
        'Tradesman 
        ParseHex(LastRxMsg.objMessage, 10, 2, Temp)
        DCInfo.Door1Options.TradesMan = CByte(Temp)
        DCInfo.Door2Options.TradesMan = CByte(Temp)

        'Door2
        'Lock Time
        ParseHex(LastRxMsg.objMessage, 14, 2, Temp)
        DCInfo.Door2Options.LockTime = CByte(Temp)
        'Open Time
        ParseHex(LastRxMsg.objMessage, 16, 2, Temp)
        DCInfo.Door2Options.OpenTime = CByte(Temp)
        'Time Profile
        ParseHex(LastRxMsg.objMessage, 18, 2, Temp)
        DCInfo.Door2Options.TimeProfile = CByte(Temp)
    End Sub
#End Region

#Region "Misc"
    Public Sub Dispose()
        'Closes the Door controller's RS232 object
        If Me.CurrentState = ParserState.Import Then
            bCloseDown = True                               'Set flag to close if not idle
        Else
            Port.Close()                                    'Shutdown Immediately
        End If

    End Sub
#End Region
End Class
