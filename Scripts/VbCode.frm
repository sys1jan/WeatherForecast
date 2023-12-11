VERSION 5.00

Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"

Begin VB.Form frmMain 

   BorderStyle     =   1  'Fixed Single

   Caption         =   "SIP"

   ClientHeight    =   2520

   ClientLeft      =   45

   ClientTop       =   330

   ClientWidth     =   5565

   BeginProperty Font 

      Name            =   "Tahoma"

      Size            =   8.25

      Charset         =   0

      Weight          =   400

      Underline       =   0   'False

      Italic          =   0   'False

      Strikethrough   =   0   'False

   EndProperty

   Icon            =   "frmMain.frx":0000

   KeyPreview      =   -1  'True

   LinkTopic       =   "Form1"

   MaxButton       =   0   'False

   ScaleHeight     =   2520

   ScaleWidth      =   5565

   StartUpPosition =   2  'CenterScreen

   Begin VB.Timer tmrSendMessages 

      Enabled         =   0   'False

      Interval        =   5000

      Left            =   4560

      Top             =   600

   End

   Begin VB.Timer tmrMessages 

      Enabled         =   0   'False

      Interval        =   250

      Left            =   5040

      Top             =   600

   End

   Begin VB.Timer tmrConnectionTest 

      Enabled         =   0   'False

      Interval        =   60000

      Left            =   5040

      Top             =   120

   End

   Begin VB.Timer tmrDiscrepancy 

      Enabled         =   0   'False

      Interval        =   60000

      Left            =   4560

      Top             =   120

   End

   Begin MSComctlLib.ListView lvwMessages 

      Height          =   1575

      Left            =   240

      TabIndex        =   0

      Top             =   360

      Width           =   4095

      _ExtentX        =   7223

      _ExtentY        =   2778

      SortKey         =   1

      View            =   3

      LabelEdit       =   1

      SortOrder       =   -1  'True

      Sorted          =   -1  'True

      LabelWrap       =   0   'False

      HideSelection   =   -1  'True

      FullRowSelect   =   -1  'True

      _Version        =   393217

      ForeColor       =   -2147483640

      BackColor       =   -2147483643

      BorderStyle     =   1

      Appearance      =   0

      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 

         Name            =   "Tahoma"

         Size            =   8.25

         Charset         =   0

         Weight          =   400

         Underline       =   0   'False

         Italic          =   0   'False

         Strikethrough   =   0   'False

      EndProperty

      NumItems        =   6

      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 

         Object.Width           =   0

      EndProperty

      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 

         Alignment       =   2

         SubItemIndex    =   1

         Text            =   "Received"

         Object.Width           =   2910

      EndProperty

      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 

         Alignment       =   2

         SubItemIndex    =   2

         Text            =   "ID"

         Object.Width           =   1058

      EndProperty

      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 

         Alignment       =   2

         SubItemIndex    =   3

         Text            =   "Status"

         Object.Width           =   2540

      EndProperty

      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 

         SubItemIndex    =   4

         Text            =   "Message"

         Object.Width           =   17780

      EndProperty

      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 

         SubItemIndex    =   5

         Text            =   "Color"

         Object.Width           =   0

      EndProperty

   End

   Begin VB.Label lblRecording 

      ForeColor       =   &H000000FF&

      Height          =   255

      Left            =   11040

      TabIndex        =   1

      Top             =   0

      Width           =   135

   End

End

Attribute VB_Name = "frmMain"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = False

Option Explicit
 
Private Const U_TIME_INTERVAL As Long = 9
 
Private m_oMessageHandler As HFCSMessageHandler.clsClient

Attribute m_oMessageHandler.VB_VarHelpID = -1

Private WithEvents m_oMessages As HFCSMessageHandler.clsNotify

Attribute m_oMessages.VB_VarHelpID = -1

Private WithEvents m_oMessagesTPLDMSGS As HFCSMessageHandler.clsNotify

Attribute m_oMessagesTPLDMSGS.VB_VarHelpID = -1

Private m_oMsgSIPConsole As HFCSMessageHandler.MsgSIPConsole

Private m_oMessageCol As clsMessageCollection
 
Private WithEvents m_oMessageHub As HFCSCMEWrapper.clsMHub

Attribute m_oMessageHub.VB_VarHelpID = -1

Private m_oCommon As HFCSSIPCommon.clsFormatMessage

Private m_oDatabase As clsDB
 
Private m_dtLastMessage As Date

Private m_lConnectionTest As Long

Private m_lPLDDiscrepFreqInMinutes As Long
 
Private m_oEventlog As HFErrorObject.clsEventLogs
 
Private Const U_MSG_MAX As Long = 500
 
Private bFlagChecked As Boolean

Private IsLineupTranslateOn As Boolean

' Color Constants

Public Enum MyColors

  U_SHADE_OF_BLUE = &HFF6633

  U_SHADE_OF_GREEN = &H9900&

  U_SHADE_OF_YELLOW = &H99CC&

  U_SHADE_OF_PURPLE = &H990099

End Enum
 
Private Sub Form_Load()

Dim lReturn As Long

Dim oPM As HFCSRemoteNotifyClient.clsRemoteLog
 
 
  On Error GoTo Errorhandler

  Set m_oEventlog = New HFErrorObject.clsEventLogs

  Set oErrorObject = New HFErrorObject.clsErrorHandling

  Set oProcessInfo = New HFErrorObject.clsProcessInformation
 
  

  Set m_oDatabase = New clsDB

  If Not m_oDatabase.SIPEnabled Then

    Set oErrorObject = Nothing

    Set oProcessInfo = Nothing

    Set m_oDatabase = Nothing

    Exit Sub

  End If

  CheckTranslateFlag

  GetFreqInMinutes

  Set m_oCommon = New HFCSSIPCommon.clsFormatMessage

  Set m_oMessageCol = New clsMessageCollection

  Set m_oMsgSIPConsole = New HFCSMessageHandler.MsgSIPConsole

  Set m_oMessageHub = New HFCSCMEWrapper.clsMHub

  Set m_oMessageHandler = g_oMessageExclusive.GetHandler()

  lReturn = m_oMessageHub.InitMessageService(U_MHUB_SIP)

  If lReturn Then

    DisplayStatus Now, 0, "InitMessageService", CStr(lReturn), vbRed

  Else

    lReturn = m_oMessageHub.AddSubscription(U_MHUB_START_LOAD)

    If lReturn Then

      DisplayStatus Now, 0, "AddSubscription(U_MHUB_START_LOAD)", CStr(lReturn), vbRed

    End If

    lReturn = m_oMessageHub.AddSubscription(U_MHUB_CLOSE_LOAD)

    If lReturn Then

      DisplayStatus Now, 0, "AddSubscription(U_MHUB_CLOSE_LOAD)", CStr(lReturn), vbRed

    End If

    lReturn = m_oMessageHub.AddSubscription(U_MHUB_TRANSFER_ULD)

    If lReturn Then

      DisplayStatus Now, 0, "AddSubscription(U_MHUB_TRANSFER_ULD)", CStr(lReturn), vbRed

    End If

    lReturn = m_oMessageHub.AddSubscription(U_MHUB_UPDATE_ULD)

    If lReturn Then

      DisplayStatus Now, 0, "AddSubscription(U_MHUB_UPDATE_ULD)", CStr(lReturn), vbRed

    End If

    lReturn = m_oMessageHub.AddSubscription(U_MHUB_TEST_MESSAGE)

    If lReturn Then

      DisplayStatus Now, 0, "AddSubscription(U_MHUB_TEST_MESSAGE)", CStr(lReturn), vbRed

    End If

    lReturn = m_oMessageHub.Subscribe()

    If lReturn Then

      DisplayStatus Now, 0, "Subscribe", CStr(lReturn), vbRed

    End If

    m_oMessageHandler.AddSubscription MSGID_SIP_CONSOLE

    Set m_oMessages = m_oMessageHandler.Receive()

    m_oMessageHandler.AddSubscription MSGID_TPLDMSGS

    Set m_oMessagesTPLDMSGS = m_oMessageHandler.Receive()

    tmrConnectionTest.Enabled = True

    tmrSendMessages.Enabled = True

    'JLW8PGC Added Startup message to display to user

    If lReturn = "0" Then

        DisplayStatus Now, 0, "SIP Service Started", "Startup", U_SHADE_OF_GREEN

    End If

  End If

  m_dtLastMessage = #1/1/1900#
 
  oProcessInfo.update_error_object Me, "Form_Load"

  oErrorObject.error_routine m_oEventlog.FeederShell, 0, LoadResString(112), oProcessInfo, OTHER_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

  Set oPM = New HFCSRemoteNotifyClient.clsRemoteLog

  oPM.LogIt U_SIP, 2, U_AUDIT, "", LoadResString(132), g_sLocalHub

  Set oPM = Nothing

Exit Sub

Errorhandler:

  oProcessInfo.update_error_object Me, "Form_Load"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

End Sub
 
Private Sub Form_QueryUnload(ByRef iCancel As Integer, ByRef iUnloadMode As Integer)

Dim sMessage As String
 
  Select Case iUnloadMode

    Case vbFormControlMenu

      sMessage = LoadResString(114)

    Case vbFormCode

      sMessage = LoadResString(115)

    Case vbAppWindows

      sMessage = LoadResString(116)

    Case vbAppTaskManager

      sMessage = LoadResString(117)

  End Select

  oProcessInfo.update_error_object Me, "Form_QueryUnload"

  oErrorObject.error_routine m_oEventlog.FeederShell, 0, sMessage, oProcessInfo, OTHER_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

End Sub
 
Private Sub Form_Unload(ByRef iCancel As Integer)

  On Error Resume Next

  Set m_oMessageHub = Nothing

  Set m_oDatabase = Nothing

  Set m_oCommon = Nothing

  Set oProcessInfo = Nothing

  Set oErrorObject = Nothing

  Set m_oEventlog = Nothing

  Set m_oMsgSIPConsole = Nothing

  Set m_oMessageHandler = Nothing

End Sub
 
Public Function TranslateBayToPosition(sBayNr As String) As String

Dim m_sSQL As String

Dim oBay As ADODB.Recordset

Dim oDatabase As HFCSDatabaseConnection.DBConn
 
  m_sSQL = "SELECT POS_NR FROM TLINUPT WHERE BAY_NR = '" & sBayNr & "'"

  Set oDatabase = New HFCSDatabaseConnection.DBConn

  Set oBay = oDatabase.Retrieve_Recordset(m_sSQL)

  If oBay.RecordCount = 1 Then

    TranslateBayToPosition = oBay.Fields("POS_NR").Value

  Else

    TranslateBayToPosition = sBayNr

  End If

  oBay.Close

  Set oBay = Nothing

  Set oDatabase = Nothing

End Function
 
Public Function TranslateIncomingPositionToBay(sPosNr As String, sBeltNa As String) As String

Dim m_sSQL As String

Dim oPositionID As ADODB.Recordset

Dim oDatabase As HFCSDatabaseConnection.DBConn
 
  Set oDatabase = New HFCSDatabaseConnection.DBConn

  m_sSQL = "SELECT BAY_NR FROM TLINUPT WHERE POS_NR = '" & sPosNr & _

           "' AND BLT_NA = '" & sBeltNa & "'"


  Set oPositionID = oDatabase.Retrieve_Recordset(m_sSQL)

  If oPositionID.RecordCount = 1 Then

    TranslateIncomingPositionToBay = oPositionID.Fields("BAY_NR").Value

  ElseIf oPositionID.RecordCount > 1 Then

    TranslateIncomingPositionToBay = sPosNr

  Else

    TranslateIncomingPositionToBay = ""

  End If

  oPositionID.Close

  Set oPositionID = Nothing

  Set oDatabase = Nothing


End Function
 
Public Sub CheckTranslateFlag()

Dim sSQL As String

Dim rsFlag As ADODB.Recordset

Dim sFlag As String

Dim oDB As HFCSDatabaseConnection.DBConn
 
On Error GoTo Errorhandler:
 
  sSQL = "SELECT VLU_IR FROM TSITCFG WHERE CFG_PAR_NR = 750"

  Set oDB = New HFCSDatabaseConnection.DBConn

  Set rsFlag = oDB.Retrieve_Recordset(sSQL)

  If rsFlag.RecordCount = 1 Then

    sFlag = rsFlag.Fields("VLU_IR").Value

    sFlag = UCase$(Trim$(sFlag))

    If sFlag = "YES" Then

      IsLineupTranslateOn = True

    Else

      IsLineupTranslateOn = False

    End If

  End If

  bFlagChecked = True

  rsFlag.Close

  Set oDB = Nothing

Exit Sub

Errorhandler:

  bFlagChecked = True

  IsLineupTranslateOn = False

  If rsFlag Is Nothing = False Then Set rsFlag = Nothing

  If oDB Is Nothing = False Then Set oDB = Nothing

End Sub
 
Private Function IsGroundContainer(messageIDType As Long, message As String) As Boolean

Dim uldTypeCode As String
 
On Error GoTo Errorhandler

  'JLW8PGC If the container type is Not 04 (Trailer) OR 07 (Package Car) OR 09 (Stack) then do not persist or process

  'Previously these were processed as 'Non-Ground' and not used causing unneeded overhead
 
    Select Case messageIDType

        Case 500, 520, 170

            'This represents the ULD-TYP-MF for all message types (ULD Type Code)

            uldTypeCode = Mid$(message, 13, 2)

        Case 160

            'This represents the TO-ULD-TYP-MF (To ULD Type Code)

            uldTypeCode = Mid$(message, 91, 2)

        Case Else

            uldTypeCode = "NonGround" 'do not process

    End Select

    '04 (Trailer) OR 07 (Package Car) OR 09 (Stack)

    If ((uldTypeCode = "04") Or (uldTypeCode = "07") Or (uldTypeCode = "09")) Then

        IsGroundContainer = True

    Else

        IsGroundContainer = False

    End If
 
Exit Function
 
Errorhandler:

IsGroundContainer = True
 
End Function
 
 
Private Sub m_oMessageHub_Message(ByVal lMessageID As Long, ByVal sMessage As String)

Dim lReturnCode As Long

Dim sPosition As String

Dim sBeltID As String

Dim sTransBay As String

Dim sNewTransBay As String

Dim lMsgLen As Long

Dim sNewPos As String

Dim sNewBeltID As String

'TT3508

Dim sBay As String
 
  On Error GoTo Errorhandler

  m_lConnectionTest = 0 ' Reset timer

  If lMessageID = HFCSSIPCommon.PLDMessageIDsEnum.START_LOAD Or lMessageID = HFCSSIPCommon.PLDMessageIDsEnum.CLOSE_LOAD _

    Or lMessageID = HFCSSIPCommon.PLDMessageIDsEnum.TRANSFER_ULD Or lMessageID = HFCSSIPCommon.PLDMessageIDsEnum.UPDATE_ULD Then

    'JLW8PGC Only proceed to IsGroundContainer Check for message types 500,520,160 and 170.

        If Not (IsGroundContainer(lMessageID, sMessage)) Then

          'JLW8PGC Only proceed it the sMessage is for a Ground Container (04 (Trailer) OR 07 (Package Car) OR 09 (Stack))

          'Previously any other type of conatiner were persisted and processed as 'Non-Ground' and not used causing unneeded overhead

          'Exit because sMessage is for a NonGround container type - (IsGroundContainer = false)

          Exit Sub

        End If

    End If

  If IsLineupTranslateOn = True Then

        If lMessageID = 500 Or lMessageID = 520 Then

            sPosition = Mid$(sMessage, 88, 4)

            sBeltID = Mid$(sMessage, 81, 7)

            sTransBay = TranslateIncomingPositionToBay(sPosition, sBeltID)

            'TT03274 - if bay is blank, do not process message

            If Len(Trim$(sTransBay)) = 0 Then

              Exit Sub

            End If

            lMsgLen = Len(sMessage)

            sMessage = Left$(sMessage, 87) & sTransBay & Mid$(sMessage, 92, lMsgLen)

        End If

        If lMessageID = 170 Then

            sPosition = Mid$(sMessage, 69, 4)

            sBeltID = Mid$(sMessage, 62, 7)

            sNewPos = Mid$(sMessage, 128, 4)

            sNewBeltID = Mid$(sMessage, 121, 7)

            sTransBay = TranslateIncomingPositionToBay(sPosition, sBeltID)

            'TT03274 - if bay is blank, do not process message

            If Len(Trim$(sTransBay)) = 0 Then

              Exit Sub

            End If

            sNewTransBay = TranslateIncomingPositionToBay(sNewPos, sNewBeltID)

            'TT03274 - if bay is blank, do not process message

            If Len(Trim$(sNewTransBay)) = 0 Then

              Exit Sub

            End If

            lMsgLen = Len(sMessage)

            sMessage = Left$(sMessage, 68) & sTransBay & Mid$(sMessage, 73, 55) & _

                       sNewTransBay & Mid$(sMessage, 132, lMsgLen)

        End If

        If lMessageID = 160 Then

          Dim iPositionBlank As Integer

          iPositionBlank = InStr(1, sMessage, " ")

          sPosition = Mid$(sMessage, iPositionBlank - 4, 4)

          sBeltID = Left$(sMessage, iPositionBlank - 5)

          sTransBay = TranslateIncomingPositionToBay(sPosition, sBeltID)

        End If

  End If

' sMessage = Replace$(sMessage, "7529", "4009")

  m_oDatabase.InsertIntoTPLDMSGS lMessageID, sMessage

  lReturnCode = m_oCommon.FromPLD(lMessageID, sMessage)

  If lReturnCode = PLDCommonErrorsEnum.OK Then

    m_oDatabase.ValidCLOB m_oCommon

    Select Case lMessageID

      Case HFCSSIPCommon.PLDMessageIDsEnum.START_LOAD

        DisplayStatus Now, lMessageID, sMessage, LoadResString(119), StatusColor(lMessageID)

        If m_oCommon.CLOB > 0 Then

          m_oDatabase.InsertByCLOB m_oCommon

          m_oDatabase.UpdateVolume m_oCommon

        Else

          m_oDatabase.InsertByLoadName m_oCommon

          'TT 3508

                    'check to see if the to-from ULDs are trailer-trailer.  If so, need to

                    'do a special database update.

                    If m_oCommon.uldTypeCode = "04" And m_oCommon.ToULDTypeCode = "04" Then

                        'both bay & Newbay in oCommon are blank, so must parse it out

                        'ULD will be "S2D1   0217 " where 0217 is the bay being transferred to

                        If IsLineupTranslateOn Then

                            sBay = sTransBay

                        Else

                            sBay = Mid$(sMessage, 88, 4)

                        End If

                        m_oDatabase.UpdateVolumeByLoadName m_oCommon, sBay

          End If

                End If

      Case HFCSSIPCommon.PLDMessageIDsEnum.TRANSFER_ULD

        DisplayStatus Now, lMessageID, sMessage, LoadResString(120), StatusColor(lMessageID)

        If m_oCommon.CLOB > 0 Then

          m_oDatabase.InsertByCLOB m_oCommon

          m_oDatabase.UpdateVolume m_oCommon ' TT03162 update volume to 1/1

        Else

          m_oDatabase.InsertByLoadName m_oCommon

          'TT 3508

                    'check to see if the to-from ULDs are stack-trailer.  If so, need to

                    'do a special database update.

                    If m_oCommon.uldTypeCode = "09" And m_oCommon.ToULDTypeCode = "04" Then

                        'both bay & Newbay in oCommon are blank, so must parse it out

                        'ULD will be "S2D1   0217 " where 0217 is the bay being transferred to

                        If IsLineupTranslateOn Then

                            sBay = sTransBay

                        Else

                            sBay = Right$(Trim$(m_oCommon.ULDNumber), 4)

                        End If

                        m_oDatabase.UpdateVolumeByLoadName m_oCommon, sBay

          End If

                End If

      Case HFCSSIPCommon.PLDMessageIDsEnum.UPDATE_ULD

        DisplayStatus Now, lMessageID, sMessage, LoadResString(121), StatusColor(lMessageID)

        If m_oCommon.CLOB > 0 Then

          m_oDatabase.InsertByCLOB m_oCommon

        Else

          m_oDatabase.InsertByLoadName m_oCommon

        End If

      Case HFCSSIPCommon.PLDMessageIDsEnum.CLOSE_LOAD

         ' if we receive a 520 with no volume, it means remove load

         DisplayStatus Now, lMessageID, sMessage, LoadResString(118), StatusColor(lMessageID)

         If m_oCommon.PackageCount = 0 Then

           If m_oCommon.CLOB > 0 Then

             m_oDatabase.DeleteByCLOB m_oCommon

           Else

             m_oDatabase.DeleteByLoadName m_oCommon

           End If

         Else

           If m_oDatabase.FindCLOB(m_oCommon.CLOB) Then

             m_oDatabase.InsertByCLOB m_oCommon

             m_oDatabase.UpdateVolume m_oCommon

           Else

             m_oDatabase.InsertByLoadName m_oCommon

           End If

         End If

      Case Else

        DisplayStatus Now, lMessageID, sMessage, LoadResString(103), MyColors.U_SHADE_OF_YELLOW

    End Select

  Else

    Select Case lReturnCode

      Case PLDCommonErrorsEnum.RETAIN_LOAD

        DisplayStatus Now, lMessageID, sMessage, LoadResString(104), MyColors.U_SHADE_OF_BLUE

      Case PLDCommonErrorsEnum.TRAINING_TRAILER

        DisplayStatus Now, lMessageID, sMessage, LoadResString(105), MyColors.U_SHADE_OF_BLUE

      Case PLDCommonErrorsEnum.NON_GROUND_CONTAINER

        DisplayStatus Now, lMessageID, sMessage, LoadResString(106), MyColors.U_SHADE_OF_BLUE

      Case PLDCommonErrorsEnum.NO_LOAD_INFORMATION

        DisplayStatus Now, lMessageID, sMessage, LoadResString(108), MyColors.U_SHADE_OF_YELLOW

      Case PLDCommonErrorsEnum.UNKNOWN_MESSAGE_TYPE

        If lMessageID = U_MHUB_TEST_MESSAGE Then

          DisplayStatus Now, lMessageID, sMessage & LoadResString(131), LoadResString(122), U_SHADE_OF_GREEN

        Else

          DisplayStatus Now, lMessageID, sMessage, LoadResString(107), MyColors.U_SHADE_OF_YELLOW

        End If

      Case Else

        DisplayStatus Now, lMessageID, sMessage, LoadResString(109), vbRed

    End Select

  End If

  tmrDiscrepancy.Enabled = False

  Exit Sub

Errorhandler:

  oProcessInfo.update_error_object Me, "m_oMessageHub_Message"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

End Sub
 
Private Function DisplayStatus(ByVal dtReceived As Date, ByVal lMessageID As Long, ByVal sMessage As String, ByVal sStatus As String, ByVal lColor As Long)

Static sSaveMessage As String

Dim li As ListItem

Dim oPM As HFCSRemoteNotifyClient.clsRemoteLog

Dim si As ListSubItem

Dim sTimeStamp As String
 
  On Error GoTo Errorhandler

  If lColor = vbRed Then

    If sSaveMessage <> sMessage Then

      sSaveMessage = sMessage

      Set oPM = New HFCSRemoteNotifyClient.clsRemoteLog

      oPM.LogIt U_SIP, 1, U_ERROR, sStatus, sMessage, g_sLocalHub

      Set oPM = Nothing

    End If

  End If

  sTimeStamp = Format(dtReceived, "MM/DD HH:MM:SS")

  If (frmMain.lvwMessages.ListItems.Count >= U_MSG_MAX) Then

    frmMain.lvwMessages.ListItems.Remove (U_MSG_MAX)

  End If

  Set li = frmMain.lvwMessages.ListItems.Add(1)

  Set si = li.ListSubItems.Add(Text:=sTimeStamp)

  si.ForeColor = lColor

  Set si = Nothing

  Set si = li.ListSubItems.Add(Text:=Format(lMessageID, "0000"))

  si.ForeColor = lColor

  Set si = Nothing

  Set si = li.ListSubItems.Add(Text:=sStatus)

  si.ForeColor = lColor

  Set si = Nothing

  Set si = li.ListSubItems.Add(Text:=sMessage)

  si.ForeColor = lColor

  Set si = Nothing

  Set si = li.ListSubItems.Add(Text:=CStr(lColor))

  Set si = Nothing

  Set li = Nothing

  m_oMsgSIPConsole.Clear

  m_oMsgSIPConsole.MessageType = U_SIP_DISPLAY_STATUS

  m_oMsgSIPConsole.Time = dtReceived

  m_oMsgSIPConsole.MsgID = lMessageID

  m_oMsgSIPConsole.Status = sStatus

    m_oMsgSIPConsole.message = sMessage

    m_oMsgSIPConsole.Color = lColor

  m_oMessageHandler.Send m_oMsgSIPConsole

  Exit Function

Errorhandler:

  oProcessInfo.update_error_object Me, "DisplayStatus"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

  Resume Next

End Function
 
'

'  Select a color for the status line

'

Private Function StatusColor(ByVal lType As Long) As Long
 
  On Error GoTo Errorhandler

  Select Case lType

    Case HFCSSIPCommon.PLDMessageIDsEnum.START_LOAD, _

         HFCSSIPCommon.PLDMessageIDsEnum.CLOSE_LOAD, _

         HFCSSIPCommon.PLDMessageIDsEnum.TRANSFER_ULD, _

         HFCSSIPCommon.PLDMessageIDsEnum.UPDATE_ULD

      StatusColor = vbBlack

    Case HFCSSIPCommon.PLDMessageIDsEnum.CANCEL_LOAD, _

         HFCSSIPCommon.PLDMessageIDsEnum.CREATE_LOAD, _

         HFCSSIPCommon.PLDMessageIDsEnum.UPDATE_LOAD, _

         HFCSSIPCommon.PLDMessageIDsEnum.DEFAULT_LINEUP, _

         HFCSSIPCommon.PLDMessageIDsEnum.CANCEL_DEFAULT_LINEUP, _

         HFCSSIPCommon.PLDMessageIDsEnum.UPDATE_DEFAULT_LINEUP

      StatusColor = U_SHADE_OF_PURPLE

    Case Else

      StatusColor = vbRed

  End Select

  Exit Function

Errorhandler:

  oProcessInfo.update_error_object Me, "StatusColor"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

  Resume Next

End Function
 
Private Sub m_oMessageHub_MessageEvent(ByVal lReturn As Long)

  On Error GoTo Errorhandler

  DisplayStatus Now, 0, Replace$(LoadResString(110), "%1", CLng(lReturn)), LoadResString(109), vbRed

  Exit Sub

Errorhandler:

  oProcessInfo.update_error_object Me, "m_oMessageHub_MessageEvent"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

  Resume Next

End Sub
 
'

'  Sometimes CIS/CME seems to not receive any messages until a message

'  is sent, so this timer will send out a message from a

'  separate EXE every hour if there's no activity.

'

Private Sub tmrConnectionTest_Timer()
 
  On Error GoTo Errorhandler

  m_lConnectionTest = m_lConnectionTest + 1

  If m_lConnectionTest >= 30 And m_lConnectionTest <= 32 Then ' 30 minutes

    ' Kill any locked up CISUTEST programs

    ' WR00354 - kill is changed to TASKKILL /F /IM

    'Shell "KILL CISUTEST*", vbHide

    Shell "TASKKILL /F /IM CISUTEST*", vbHide

  ElseIf m_lConnectionTest > 60 Then ' 1 Hour

    SendTestMessage

    m_lConnectionTest = 0

  End If

  Exit Sub

Errorhandler:

  oProcessInfo.update_error_object Me, "tmrConnectionTest_Timer"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

  Resume Next

End Sub
 
Private Sub tmrDiscrepancy_Timer()

'Dim lDiscrepancies As Long

'Dim oIPLD As IPLD.clsComm

'Dim oScratch As HFCSScratchPadClient.clsScratchPad

'Dim sMessage As String

'

'  On Error GoTo Errorhandler

'  If DateDiff("n", m_dtLastMessage, Now) >= m_lPLDDiscrepFreqInMinutes Then

'    Set oIPLD = New IPLD.clsComm

'    lDiscrepancies = oIPLD.NumberOfDiscrepancies(m_oDatabase.GetConnection)

'    Set oIPLD = Nothing

'    If lDiscrepancies > 0 Then

'      If lDiscrepancies > 1 Then

'        sMessage = Replace$(LoadResString(125), "%1", CStr(lDiscrepancies))

'      Else

'        sMessage = LoadResString(126)

'      End If

'      sMessage = sMessage & " " & LoadResString(101)

'      m_dtLastMessage = Now

'      Set oScratch = New HFCSScratchPadClient.clsScratchPad

'      oScratch.AddScratchPadRecord GetPlanNumber(), sMessage, False, ScratchPadMsgTypes.IPLD, NonInitMsg, "", BothScratchPads, 1024, sMessage

'      Set oScratch = Nothing

'    Else

'      m_dtLastMessage = #1/1/1900#

'      tmrDiscrepancy.Enabled = False

'    End If

'    Set oIPLD = Nothing

'  End If

'  Exit Sub

'

'Errorhandler:

'  oProcessInfo.update_error_object Me, "tmrDiscrepancy_Timer"

'  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

'  Set oIPLD = Nothing

'  Set oScratch = Nothing

End Sub
 
Private Function GetPlanNumber() As Long

Dim oPlanNumber As HFCSFlowControlMessages.clsMessage
 
  On Error GoTo Errorhandler

  Set oPlanNumber = New HFCSFlowControlMessages.clsMessage

  GetPlanNumber = oPlanNumber.FlowControlActive

  Set oPlanNumber = Nothing

  Exit Function

Errorhandler:

  Set oPlanNumber = Nothing

  GetPlanNumber = 0 ' No plan number
 
End Function
 
Private Sub m_oMessages_Receive(ByVal oMessage As Object)

  On Error GoTo Errorhandler

  m_oMessageCol.Push oMessage

  tmrMessages.Enabled = True

  Set oMessage = Nothing

  Exit Sub

Errorhandler:

  oProcessInfo.update_error_object Me, "m_oMessages_Receive"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

  Resume Next

End Sub
 
Private Sub tmrMessages_Timer()

Dim oMessage As HFCSMessageHandler.MsgSIPConsole
 
  On Error GoTo Errorhandler

  tmrMessages.Enabled = False

  Do While (m_oMessageCol.Count > 0)

    Set oMessage = m_oMessageCol.Pop()

    Select Case oMessage.MessageType

      Case HFCSMessageHandler.SIPMsgEnum.U_SIP_PERFORM_TEST

        SendTestMessage

      Case HFCSMessageHandler.SIPMsgEnum.U_SIP_REQUEST_REFRESH

        RefreshConsole

    End Select

    Set oMessage = Nothing

  Loop

  Exit Sub

Errorhandler:

  oProcessInfo.update_error_object Me, "tmrMessages_Timer"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

  Resume Next

End Sub
 
Private Sub m_oMessagesTPLDMSGS_Receive(ByVal oMessage As Object)

  On Error GoTo Errorhandler

  tmrSendMessages.Enabled = True

  Set oMessage = Nothing

  Exit Sub

Errorhandler:

  oProcessInfo.update_error_object Me, "m_oMessagesTPLDMSGS_Receive"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

  Resume Next

End Sub
 
Private Sub tmrSendMessages_Timer()

Dim lReturn As Long

Dim rsMessages As ADODB.Recordset

Dim sType As String

Dim oScratch As HFCSScratchPadClient.clsScratchPad
 
  On Error GoTo Errorhandler

  tmrSendMessages.Enabled = False

  Set rsMessages = m_oDatabase.GetRecordsToSend()

  If Not rsMessages Is Nothing Then

    Do While Not rsMessages.EOF

      lReturn = m_oMessageHub.Publish(rsMessages.Fields("MSG_TE").Value, rsMessages.Fields("MSG_ID").Value)

      If lReturn = 0 Then

        Select Case rsMessages.Fields("MSG_ID").Value

          Case HFCSSIPCommon.PLDMessageIDsEnum.CANCEL_LOAD

            sType = LoadResString(127)

          Case HFCSSIPCommon.PLDMessageIDsEnum.CREATE_LOAD

            sType = LoadResString(128)

          Case HFCSSIPCommon.PLDMessageIDsEnum.UPDATE_LOAD

            sType = LoadResString(121)

          Case HFCSSIPCommon.PLDMessageIDsEnum.DEFAULT_LINEUP

            sType = LoadResString(129)

          Case HFCSSIPCommon.PLDMessageIDsEnum.CANCEL_DEFAULT_LINEUP

            sType = LoadResString(130)

          Case HFCSSIPCommon.PLDMessageIDsEnum.UPDATE_DEFAULT_LINEUP

            sType = LoadResString(133)

        End Select

        DisplayStatus Now, rsMessages.Fields("MSG_ID").Value, rsMessages.Fields("MSG_TE").Value, sType, StatusColor(rsMessages.Fields("MSG_ID").Value)

        m_oDatabase.FlagRecordAsSent rsMessages.Fields("STR_MSG_TS").Value, rsMessages.Fields("MSG_ID").Value

      Else

        DisplayStatus Now, 0, rsMessages.Fields("MSG_TE").Value, CStr(lReturn), vbRed

        Set oScratch = New HFCSScratchPadClient.clsScratchPad

        oScratch.AddScratchPadRecord GetPlanNumber(), "", False, ScratchPadMsgTypes.IPLD, NonInitMsg, "", FeederScratchPad, 1091

        Set oScratch = Nothing

      End If

      rsMessages.MoveNext

    Loop

    m_oDatabase.GetConnection.CloseRecordSet rsMessages

  End If

  Exit Sub

Errorhandler:

  oProcessInfo.update_error_object Me, "tmrSendMessages_Timer"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

  Set oScratch = Nothing

End Sub
 
Private Sub SendTestMessage()

Dim lRc As Long

Dim oPM As HFCSRemoteNotifyClient.clsRemoteLog
 
  On Error GoTo Errorhandler

  lRc = ShellAndWait("CISUTEST /P " & Right$("0000" & CStr(U_MHUB_TEST_MESSAGE), 4), 15000, vbHide, AbandonWait)

  If lRc <> 0 Then

    Set oPM = New HFCSRemoteNotifyClient.clsRemoteLog

    oPM.LogIt U_SIP, 1, U_WARNING, CStr(lRc), LoadResString(134), g_sLocalHub

    Set oPM = Nothing

    Shell "TASKKILL /F /IM HFCSSIPCommon.exe", vbHide

    Shell "TASKKILL /F /IM HFCSSIP.exe", vbHide

    Shell "TASKKILL /F /IM HFCSPCMFComm.exe", vbHide

  End If

  Exit Sub

Errorhandler:

  oProcessInfo.update_error_object Me, "SendTestMessage"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

  Resume Next

End Sub
 
Private Sub GetFreqInMinutes()

Dim lTemp As Long

Dim oDB As HFCSDatabaseConnection.DBConn

Dim oConfig As HFCSGeneralConfig.GeneralConfigs

Dim tConfigs() As HFCSGeneralConfig.ConfigurationInfo

  On Error GoTo Errorhandler

  Set oDB = New HFCSDatabaseConnection.DBConn

  oDB.Popups = False

  oDB.LocalHub = g_sLocalHub

  Set oConfig = New HFCSGeneralConfig.GeneralConfigs

  tConfigs = oConfig.GetConfigurationInfo(CONFIG_AS_ARRAY, oDB, 910)

  If UBound(tConfigs) = 0 Then

    m_lPLDDiscrepFreqInMinutes = U_TIME_INTERVAL

  Else

    lTemp = CLng(tConfigs(1).NumericValue)

    If (lTemp <= 0) Then

      m_lPLDDiscrepFreqInMinutes = U_TIME_INTERVAL

    ElseIf (lTemp > 1440) Then

      m_lPLDDiscrepFreqInMinutes = 1440 ' Once a day

    Else

      m_lPLDDiscrepFreqInMinutes = lTemp

    End If

  End If

  Set oConfig = Nothing

  Set oDB = Nothing

  Exit Sub

Errorhandler:

  m_lPLDDiscrepFreqInMinutes = U_TIME_INTERVAL

  oProcessInfo.update_error_object Me, "GetFreqInMinutes"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

End Sub
 
Private Sub RefreshConsole()

Dim lLoop As Long
 
  On Error GoTo Errorhandler

  If (lvwMessages.ListItems.Count > 0) Then

    m_oMsgSIPConsole.Clear

    m_oMsgSIPConsole.MessageType = U_SIP_DISPLAY_STATUS

    m_oMessageHandler.Send m_oMsgSIPConsole

    For lLoop = 1 To lvwMessages.ListItems.Count

      m_oMsgSIPConsole.Clear

      m_oMsgSIPConsole.MessageType = U_SIP_DISPLAY_STATUS

      m_oMsgSIPConsole.Time = lvwMessages.ListItems.Item(lLoop).SubItems(1)

      m_oMsgSIPConsole.MsgID = lvwMessages.ListItems.Item(lLoop).SubItems(2)

      m_oMsgSIPConsole.Status = lvwMessages.ListItems.Item(lLoop).SubItems(3)

            m_oMsgSIPConsole.message = lvwMessages.ListItems.Item(lLoop).SubItems(4)

            m_oMsgSIPConsole.Color = CLng(lvwMessages.ListItems.Item(lLoop).SubItems(5))

      m_oMessageHandler.Send m_oMsgSIPConsole

    Next lLoop

  End If

  Exit Sub

Errorhandler:

  oProcessInfo.update_error_object Me, "RefreshConsole"

  oErrorObject.error_routine m_oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_SIP, NO_POP_UP, , , g_sLocalHub

End Sub
