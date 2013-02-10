
Option Explicit




Private m_Caller As IIDCCaller

Public Event ClearRequested()
Public Event BalanceRequested()

Private Const cError = "Error"

Private WithEvents mMenuControl As CMenuControl 'Centralisation
Attribute mMenuControl.VB_VarHelpID = -1

Private Function GetCustCon() As IIDCCUSTOMERCONTEXTLib.IIDCCustomerContext
    If m_Caller Is Nothing Then Exit Function
    Set GetCustCon = m_Caller.Parent
End Function

Private Function IsHidden() As Boolean
    If m_Caller Is Nothing Then Exit Function
    IsHidden = m_Caller.Parent.Hidden
End Function


Friend Sub ChangeStatus(ByVal NewStatus As IdAndVStatusMask)
    DetermineBalanceEnabled
    If NewStatus <> idvIdNone Then SetCustomerIndicators
End Sub


Private Sub cmbCustomers_Click()
    On Error GoTo Handler
    
    Set m_Caller.SelectedCustomer = m_Caller.Customers(cmbCustomers.ListIndex + 1)

    ' Centralisation: only for ConsiderAsLive customers.
    ' NB - why isn't this in the Property Let for SelectedCustomer?
    ' Premier release - may now process (some) suspended callers.
    If m_Caller.Customer.ConsiderAsLive And m_Caller.SelectedCustomer.CisId <> m_Caller.CisId Then
        Dim CustStatus As String
        CustStatus = m_Caller.Customers(cmbCustomers.ListIndex + 1).BcallCustomer.PersonalCustomerDetails(1).OperationalCode
        If CustStatus <> "L" And CustStatus <> "N" Then
            Select Case CustStatus
            Case "T"
                CustStatus = "Targeted"
            Case "S"
                CustStatus = "Suspended"
            Case "R"
                CustStatus = "Responded"
            Case "A"
                CustStatus = "Acknowledged"
            Case "W"
                CustStatus = "Welcome"
            Case "C"
                CustStatus = "Closed"
            Case "P"
                CustStatus = "Pending"
            End Select
    
            If Not m_Caller.AllowServicingAfterSuspension Then
                IDCContainer.IDCLog.ErrorDialog "Status of requested customer is " & CustStatus & " - Unable to service", idcErrorActionCancel
                
                ' Now switch back to the first customer on the list (should already be read from mainframe).
                
                If m_Caller.Customers(1).BcallCustomer.PersonalCustomerDetails.Count <> 1 Then
                    IDCContainer.IDCLog.Record "cmbCustomers_Click: No PersonalCustomerDetails(1)" & vbCrLf & m_Caller.Customers(1).BcallCustomer.Dump, idcSeverityError, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
                    RaiseEvent ClearRequested
                ElseIf m_Caller.Customers(1).BcallCustomer.PersonalCustomerDetails(1).OperationalCode <> "L" And m_Caller.Customers(1).BcallCustomer.PersonalCustomerDetails(1).OperationalCode <> "N" Then
                    IDCContainer.IDCLog.Record "cmbCustomers_Click: First customer is not live:" & vbCrLf & m_Caller.Customers(1).BcallCustomer.Dump, idcSeverityError, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
                    RaiseEvent ClearRequested
                Else
                    Set m_Caller.SelectedCustomer = m_Caller.Customers(1)
                End If
            End If
        End If
    End If
    
    Exit Sub
Handler:
    IDCContainer.IDCLog.Record "cmbCustomers_Click: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
    
    If m_Caller.Customers(cmbCustomers.ListIndex + 1).BcallCustomer.CustomerNotFound Then
        IDCContainer.IDCLog.ErrorDialog "Customer not known", idcErrorActionCancel
    Else
        If IDCContainer.IDCLog.ErrorDialog("Can not switch to the selected customer" & vbCrLf & Err.Description, idcErrorActionCancel Or idcErrorActionRetry) = idcErrorActionRetry Then
            Resume
        End If
    End If
    
    Dim Cust As IIDCCustomer
    Dim Index As Long
    
    For Each Cust In m_Caller.Customers
        If Cust.CisId = m_Caller.SelectedCustomer.CisId Then Exit For
        Index = Index + 1
    Next
    
    Debug.Assert Not (Cust Is Nothing)
    
    cmbCustomers.ListIndex = Index
End Sub

'Private Sub cmbCustomers_GotFocus()
'    Dim x As Integer: x = SendMessage(cmbCustomers.hWnd, CB_SHOWDROPDOWN, 0, 0&)
'End Sub

Private Sub cmdBalance_Click()
On Error GoTo ErrHandler
    BalanceEnquiry
Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "cmdBalance_Click: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Private Sub cmdClear_Click()
    On Error GoTo ErrHandler
    Dim bcxData As GlobalBCX1: Set bcxData = IDCContainer.IDCServiceManager.Retrieve("BCXGlobalData")
    If bcxData.prConfirm("Customer Call Completed") = "Y" Then
        ' ActionPrompts: No longer set this reference to Nothing, as the QueryDestroyData could cancel
        ' the clear
        'Set m_Caller = Nothing
        ' SIR 2385: SetFocus BEFORE calling Clear, as this will start ID&V if there is a call waiting.
        IDCContainer.IDCMDIWindow.SetFocus
        RaiseEvent ClearRequested
    End If
    Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "cmdClear_Click: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Private Sub cmdDetail_Click() 'Centralisation
    DisplayDetail
End Sub

Private Sub cmdEntries_Click() 'Centralisation
    If mMenuControl Is Nothing Then Set mMenuControl = IDCContainer.IDCServiceManager.Retrieve(MENU_CONTROL)
    mMenuControl.GetAuthorisation ACCOUNT_ENTRIES_AUTHORISATION_ID
End Sub

Private Sub cmdLetters_Click()
    Set mMenuControl = IDCContainer.IDCServiceManager.Retrieve(MENU_CONTROL)
    mMenuControl.GetAuthorisation CUSTOMER_LETTERS
End Sub

Private Sub cmdMiscSummary_Click()
    frmMiscBalance.Visible = False
    frmMiscBalance.Enabled = False
    frmSummaryBalance.Visible = True
    frmSummaryBalance.Enabled = True
End Sub

Private Sub cmdMore_Click()
    If mMenuControl Is Nothing Then Set mMenuControl = IDCContainer.IDCServiceManager.Retrieve(MENU_CONTROL)
    mMenuControl.GetAuthorisation PRODUCT_DETAILS_ID
End Sub

Private Sub cmdOutStandingTrans_Click() 'Centralisation
    If mMenuControl Is Nothing Then Set mMenuControl = IDCContainer.IDCServiceManager.Retrieve(MENU_CONTROL)
    mMenuControl.GetAuthorisation OUTSTANDING_TRANS_AUTHORISATION_ID
End Sub

Private Sub cmdLoanEntries_Click() 'Centralisation
    If mMenuControl Is Nothing Then Set mMenuControl = IDCContainer.IDCServiceManager.Retrieve(MENU_CONTROL)
    mMenuControl.GetAuthorisation ACCOUNT_ENTRIES_AUTHORISATION_ID
End Sub

' IDT25 P0127495 - July 2002 - Correct tab sequence
Private Sub fpList1_GotFocus()
    If SSTab1.Tab = 2 Then
        cmdLetters.SetFocus
    End If
End Sub

' IDT25 P0127495 - July 2002 - Correct tab sequence
Private Sub cmdBalance_GotFocus()
    If SSTab1.Tab = 2 Then
        If cmdDiaryEntry.Enabled Then
            cmdDiaryEntry.SetFocus
        Else
            SSTab1.SetFocus
        End If
    End If
End Sub

Private Sub indCustomerIndicators_CmdCareClicked()
    '
    Dim MenuControl As IMENUCONTROLLib.CMenuControl
    Set MenuControl = IDCContainer.IDCServiceManager.Retrieve("MenuControl")
    
    If MenuControl.IsServiceEnabled(SETCUSTMARKERS) Then
        MenuControl.ActivateMenuItem SETCUSTMARKERS
    End If
End Sub

Private Sub mMenuControl_AuthStatus(ByVal Task As String, ByVal IsAuthorised As Boolean) 'Centralisation
    Set mMenuControl = Nothing
    
    If IsAuthorised Then
        RunTask Task
    End If
End Sub

Private Sub RunTask(Task As String) 'Centralisation

    Dim Service As IIDCServiceManagement
    Dim EnquiryObject As IRunEnquiry
    
    Dim AlreadyRunning As Boolean
    
    Dim strMenu As String
    Dim strNonMenu As String
    
    If StrComp(Task, OUTSTANDING_TRANS_AUTHORISATION_ID, vbTextCompare) = 0 Then
        strMenu = OUTSTANDING_TRANS_MENU
        strNonMenu = OUTSTANDING_TRANS_NON_MENU
    ElseIf StrComp(Task, ACCOUNT_ENTRIES_AUTHORISATION_ID, vbTextCompare) = 0 Then
        strMenu = ACCOUNT_ENTRIES_MENU
        strNonMenu = ACCOUNT_ENTRIES_NON_MENU
    ElseIf StrComp(Task, PRODUCT_DETAILS_ID, vbTextCompare) = 0 Then
        strMenu = PRODUCT_DETAILS
        strNonMenu = ""         ' There is no special non-menu version.
    ElseIf StrComp(Task, CUSTOMER_LETTERS, vbTextCompare) = 0 Then
        strMenu = CUSTOMER_LETTERS
        strNonMenu = ""         ' There is no special non-menu version.
    Else
        ' IDT Release 24
        ' The most likely cause of this is a programming error, as something has set
        ' mMenuControl without making sure it is unset after use. However it is also possible
        ' (but very unlikely) that some other service may have requested a GetAuthorisation while
        ' we are waiting for our authorisation to return.
        
        IDCContainer.IDCLog.Trace "RunTask(""" & Task & """) unexpected" & vbCrLf & IDCContainer.IDCLog.StackDump()
        Debug.Assert False      ' If you stop here, there is probably something wrong!
        
        Exit Sub
    End If
    
    Set Service = IDCContainer.IDCServiceManager.FindInstanceByName(strMenu)
    If Service Is Nothing Then
        If strNonMenu <> "" Then
            Set Service = IDCContainer.IDCServiceManager.FindInstanceByName(strNonMenu)
            If Service Is Nothing Then
                Set Service = IDCContainer.IDCServiceManager.Retrieve(strNonMenu)
            Else
                AlreadyRunning = True
            End If
        Else
            Set Service = IDCContainer.IDCServiceManager.Retrieve(strMenu)
        End If
    Else
        AlreadyRunning = True
    End If

    If Service Is Nothing Then
        IDCContainer.IDCLog.Record "Could not retrieve " & strMenu & " Service", idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
        IDCContainer.IDCLog.ErrorDialog "Could not display " & strMenu & " Service", idcErrorActionCancel
    ElseIf StrComp(Task, PRODUCT_DETAILS_ID, vbTextCompare) = 0 Then
        Dim DispLoanDetails As IDisplayLoanDetailsLib.IDisplayLoanDetails
        Set DispLoanDetails = Service
        DispLoanDetails.DisplayAccountDetails m_Caller.SelectedAccount
    ElseIf StrComp(Task, CUSTOMER_LETTERS, vbTextCompare) = 0 Then
        ' The Letters service has been written in a different way to the others. We must not call RunEnquiry if
        ' we have started the service (i.e. it was not already running).
        If AlreadyRunning Then
            Set EnquiryObject = Service
            EnquiryObject.RunEnquiry m_Caller.SelectedCustomer
        End If
    Else
        Set EnquiryObject = Service
        EnquiryObject.RunEnquiry m_Caller.SelectedAccount
    End If
End Sub

Private Sub cmdLoanSummary_Click() 'Centralisation
    frmLoanBalance.Visible = False
    frmLoanBalance.Enabled = False
    frmSummaryBalance.Visible = True
    frmSummaryBalance.Enabled = True
End Sub

Private Sub cmdDiaryEntry_Click()
    DiaryEntry
End Sub

Private Sub cmdBalanceSummary_Click()
    frmDetailBalance.Visible = False
    frmDetailBalance.Enabled = False
    frmSummaryBalance.Visible = True
    frmSummaryBalance.Enabled = True
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    Clear
    DisableAllControls
    
Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "Form Load: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub




Private Sub fpList1_DblClick()
    BalanceEnquiry
End Sub

Private Sub fpList1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then BalanceEnquiry
End Sub

Private Sub fpList1_SelChange(ItemIndex As Long)
On Error GoTo ErrHandler

    'Debug.Print "FP SelChange " & ItemIndex

    fpList1.Row = fpList1.ListIndex
    fpList1.ColFromName = "Balance"
    txtBalance = RTrim(fpList1.ColList)
    fpList1.ColFromName = "Overdraft"
    txtOverdraft = fpList1.ColList
    fpList1.ColFromName = "Ledger"
    txtUpdatedLedger = RTrim(fpList1.ColList)
    'ADE Problem DCH 00031165 (IO302190)
    'Set m_Caller.SelectedAccount = m_Caller.SelectedCustomer.Accounts(fpList1.ListIndex + 1)
    Dim strSortCode As String
    fpList1.ColFromName = "SortCode"
    strSortCode = RTrim(fpList1.ColList)
    Dim strAccount As String
    fpList1.ColFromName = "AccountNumber"
    strAccount = RTrim(fpList1.ColList)
    
    ' OpenPlan - Use FindAccount so that Misc Accounts are searched as well
    Set m_Caller.SelectedAccount = m_Caller.SelectedCustomer.FindAccount(strSortCode, strAccount)
'    Dim objSelectedAccount As Object
'    For Each objSelectedAccount In m_Caller.SelectedCustomer.Accounts
'        If strSortCode = objSelectedAccount.Sortcode And strAccount = objSelectedAccount.AccountNumber Then
'            Set m_Caller.SelectedAccount = objSelectedAccount
'        End If
'    Next objSelectedAccount
    'ADE Problem DCH 00031165 (IO302190)
    
    'December 2000
    fpList1.ColFromName = "UnclearedAmount"
    If Trim(fpList1.ColList) <> "" Then
        Dim strUnclearedAmount As String
        Dim strClearedBalance As String
        strUnclearedAmount = fpList1.ColList
        fpList1.ColFromName = "ClearedBalance"
        strClearedBalance = fpList1.ColList
        lblWarningText.Caption = FormatWarningText(strUnclearedAmount, strClearedBalance)
        lblWarning.Visible = True
        lblWarningText.Visible = True
    Else
        fpList1.ColFromName = "IsBaProduct"
        If fpList1.ColList = CStr(False) Then
            lblWarningText.Caption = lblWarningText2.Caption
            lblWarning.Visible = True
            lblWarningText.Visible = True
        Else
            lblWarning.Visible = False
            lblWarningText.Visible = False
        End If
    End If
Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "fpList1_SelChange: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Public Sub ClearCustomerIndicators()
    With indCustomerIndicators
        .Clear
        .Enabled = False
    End With
End Sub

Public Sub SetCustomerIndicators()
    On Error GoTo ErrHandler
    
    If IsHidden() Then Exit Sub
    If m_Caller Is Nothing Then Exit Sub
    Dim CustObj As PIDCCustomerData.CCustomerData
    Set CustObj = m_Caller.Customer

    Dim CustomerBusy As Boolean
    Dim CustomerFailed As Boolean
    
    If CustObj.IsBusy Then
        CustomerBusy = True
    Else
        CustomerFailed = Not (CustObj.errorObject(ErrObj_CUSTOMER) Is Nothing)
        
        ' ADE 00036002 : Do not crash if it is a dummy.
        If Len(m_Caller.CisId) = 10 Then
            Debug.Assert (CustObj.BaseDetails.Count = 1 Or CustObj.PersonalDetails.Count = 1) Eqv (CustObj.errorObject(ErrObj_CUSTOMER) Is Nothing)
        End If
    End If
    
    IDCContainer.IDCLog.Trace "SetCustomerIndicators: CustomerBusy=" & CustomerBusy & " CustomerFailed=" & CustomerFailed
    
    With indCustomerIndicators
        
        ' =====================================================================================================
        ' Set up the "basic" indicators, that are normally read from the CPL
        .Complaint = m_Caller.Complaint
        .VIP = m_Caller.VIP
        .Staff = m_Caller.Staff
        .Care = GetSpecialCareReasons()
        
        ' =====================================================================================================
        ' The following indicators are derived from the Customer Download data. They may not yet be available,
        ' or the download might have failed.
        
        If CustomerBusy Then
            .Portfolio = stsPending
            .SegmentInfo = stsPending
            .Premier = stsPending
            .OpenPlan = stsPending
            .Limits = stsPending
        ElseIf CustomerFailed Then
            .Portfolio = stsError
            .SegmentInfo = stsError
            .Premier = stsError
            .OpenPlan = stsError
            .Limits = stsError
        Else
            ' =====================================================================================================
            ' Premier indicator. NB may also be set if the segment information indicates premier.
            ' Set the "Premium" indicator (NB: some versions of the design have this displayed
            ' on the screen, some don't.)
            
            ' ADE 00036002 : Do not crash if it is a dummy.
            If Len(m_Caller.CisId) = 10 Then
                If m_Caller.Premier Then
                    If m_Caller.OnDemand Then
                        .Premier = "Pond"
                    Else
                        .Premier = "Premier"
                    End If
                Else
                    .Premier = False
                End If
            End If
            
            ' =====================================================================================================
            ' OpenPlan indicator.
            
            Dim Acc As IIDCAccount
            .OpenPlan = stsFalse
            
            If Not m_Caller.PersonalCustomer Is Nothing Then
                For Each Acc In m_Caller.PersonalCustomer.MiscAccounts
                    If Acc.ProductReference.InGroup(idcACIOpenPlanArrangement) Then
                       .OpenPlan = stsTrue
                       Exit For
                    End If
                Next
            End If
            
            
            ' =====================================================================================================
            ' Update "Care" indicator if B&D
            .Care = GetSpecialCareReasons()
            
            ' Decode the Business Unit and Segment data.
            ' Do the Business Unit data first. This may set the "SegmentInfo" indicator to
            ' "Rel" (meaning that the customer is Relationship manageged). This may be overwritten
            ' by a more specific value when we decode the segment data.
            
            DecodeBusinessUnitData CustObj
            
            DecodeSegmentData CustObj
            
            ' =====================================================================================================
            ' Limits Indicator.
            
            Debug.Assert m_Caller.SelectedCustomer.LimitsCode <> ""
            
            DecodeLimitsCode
            
        End If
    End With

    ' If we now know the customer relationships, we may be able to enable the diary button
    DetermineDiaryEnabled
    
    
Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "SetCustomerIndicators: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

' R26 - allow multiple special care reasons (for Do Not Promote IVR).
Private Function GetSpecialCareReasons() As Collection
On Error GoTo Handler
    Dim Coll As Collection
    Set Coll = New Collection
    
    If Not m_Caller.PersonalCustomer Is Nothing Then
        If m_Caller.PersonalCustomer.BadAndDoubtful Then
            Coll.Add "Bad and Doubtful"
        End If
    End If
    
    If m_Caller.SpecialCare Then
        Coll.Add m_Caller.SpecialCareReason
    End If
    
    Dim Cust As CCustomerData
    Set Cust = m_Caller.Customer
    
    If Not Cust Is Nothing Then
        Dim ExclusionMethod As CExclusionMethod
        
        For Each ExclusionMethod In Cust.ExclusionMethods
            If ExclusionMethod.IsCurrent Then
                Dim LRefItems As IIDCLocalReferenceItems
                Set LRefItems = GetLRefItems("ExclsnCare", Trim(ExclusionMethod.ContactMethod))
                If LRefItems.Count > 0 Then
                    Debug.Assert LRefItems.Count = 1
                    
                    Coll.Add LRefItems(1).Description
                End If
            End If
        Next
    End If
    
    Set GetSpecialCareReasons = Coll
Exit Function
Handler:
      Err.Raise Err.Number, Err.Source, "GetSpecialCareReasons: " & Err.Description
End Function

Private Sub DecodeSegmentData(ByVal CustObj As PIDCCustomerData.CCustomerData)
    lstSegmentMarkers.Clear

    Dim Seg As Customer.CActionPrompt
    Dim SegDecode As CSegmentInfo
    Dim LocRef As IIDCLocalRefData
    Dim LocrefItem As IIDCLocalReferenceItem

    Dim PAM As Boolean
    Dim GraduateRelationship As Boolean
    Dim Graduate  As Boolean
    Dim Student As Boolean
    Dim OtherRelationship As Boolean
    
    For Each Seg In CustObj.ActionPrompts
        If Seg.state <> Deleted Then
            Set SegDecode = New CSegmentInfo
            With SegDecode
                .Initialise Seg
                If .OfInterest Then
                    lstSegmentMarkers.AddItem .Decode
                    
                    ' April 2001 release:
                    ' Now look to see if Premier etc. need to be set.
                    ' ADE DCH  00042917 - Do not use the SegmentInfo to light up the
                    ' Premier indicator (use the IIDCCaller.Premier property ONLY). This
                    ' is because the segment info is seen as unreliable, and we do not want a
                    ' mismatch between the displayed indicator and the IIDCCaller property.
                    ' If .Premier Then indCustomerIndicators.Premier = True
                    
                    ' ADE DCH  00035143:
                    ' Segment indicators: PAM overrides GradR, overrides Grad, overrides Student, so
                    ' just keep a note of which ones we find (the order they come back in may not
                    ' be the same as this precedence).
                    ' (In practice, there should only be one marker)
                    If .PAM Then PAM = True
                    If .GraduateRelationship Then GraduateRelationship = True
                    If .Graduate Then Graduate = True
                    If .Student Then Student = True
                    If .OtherRelationship Then OtherRelationship = True
                End If
            End With
        End If
    Next
    
    If PAM Then
        indCustomerIndicators.SegmentInfo = "PAM"
    ElseIf GraduateRelationship Then
        indCustomerIndicators.SegmentInfo = "Grad R"
    ElseIf Graduate Then
        indCustomerIndicators.SegmentInfo = "Grad"
    ElseIf Student Then
        indCustomerIndicators.SegmentInfo = "Studt"
    ElseIf OtherRelationship Then
        indCustomerIndicators.SegmentInfo = "Rel"
    End If
End Sub


Private Sub DecodeBusinessUnitData(ByVal CustObj As PIDCCustomerData.CCustomerData)
    If (CustObj.CustDataMask And CBUR_PRM) = CBUR_PRM Then
        Dim BusUnits As Customer.CBusinessUnits
        Dim BusUnit As Customer.CBusinessUnit
        Set BusUnits = CustObj.BusinessUnits
        
        indCustomerIndicators.Portfolio = stsFalse
        indCustomerIndicators.SegmentInfo = ""
        
        For Each BusUnit In BusUnits
            If BusUnit.state <> Deleted Then
                If BusUnit.IsCurrent Then
                    If BusUnit.CustomerBusinessUnitRelationTypeCode = "R" Then
                        indCustomerIndicators.SegmentInfo = "Rel"
                    End If
                    If BusUnit.CustomerBusinessUnitRelationTypeCode = "P" Then
                        indCustomerIndicators.Portfolio = True
                    End If
                End If
            End If
        Next
    End If
End Sub

Private Sub DecodeLimitsCode()
    indCustomerIndicators.Limits = ""
    
    If m_Caller.PersonalCustomer Is Nothing Then Exit Sub
    
    Debug.Assert m_Caller.PersonalCustomer.LimitsCode <> ""
    If m_Caller.PersonalCustomer.LimitsCode = "" Then Exit Sub
    
    Dim LocalRefData As IIDCLocalRefData
    Dim LocalRefItems As IIDCLocalReferenceItems
    Set LocalRefData = IDCContainer.IDCServiceManager.Retrieve("LocalRefData")
    Set LocalRefItems = LocalRefData.GetLocalRefItems("LimitsInd", Code:=m_Caller.PersonalCustomer.LimitsCode, CurrentOnly:=True)
    
    Debug.Assert LocalRefItems.Count = 1
    If LocalRefItems.Count <> 1 Then Exit Sub
    
    indCustomerIndicators.Limits = Trim(LocalRefItems(1).Description)
End Sub

Public Sub Initialise(Caller As IIDCCaller)
    On Error GoTo ErrHandler
    
    Set m_Caller = Caller

    SetupScreenHeader

    txtBalance = ""
    txtOverdraft = ""
    txtUpdatedLedger = ""
    'December 2000
    lblWarning.Visible = False
    lblWarningText.Visible = False
    
    IDCContainer.IDCLog.Trace "Customer Context - Start"
Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "CustCtx.ServiceForm.Initialise: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Public Sub BalanceEnquiry()
On Error GoTo ErrHandler
    If Not cmdBalance.Enabled Then Exit Sub
        
    txtBalance = ""
    txtOverdraft = ""
    txtUpdatedLedger = ""
    'December 2000
    lblWarning.Visible = False
    lblWarningText.Visible = False
    RaiseEvent BalanceRequested

Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "Balance Enquiry: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Public Sub DiaryEntry()
On Error GoTo ErrHandler
    If Not cmdDiaryEntry.Enabled Then Exit Sub
    
    Dim Recipient As String
    
    Dim DiaryForm As frmDiaryMsg
    Set DiaryForm = New frmDiaryMsg
    
    Dim Params As Collection
    Set Params = New Collection
    
    With lstRelationshipManagers
        If .ListIndex < 0 Then Exit Sub ' Button should never have been Enabled.
        .ColFromName = "colBusuId"
        Params.Add .ColText, Key:="BUSN"
        .ColFromName = "colEmpId"
        Params.Add .ColText, Key:="EMP"
        
        .ColFromName = "colEmpSurname"
        Recipient = .ColText
        
        .ColFromName = "colEmpInitials"
        Recipient = Recipient & " " & .ColText
        
        Params.Add "Sales", Key:="CATEGORY"
        Params.Add "Low", Key:="PRIORITY"
    End With
    
    DiaryForm.Recipient = Recipient
    DiaryForm.Show vbModal
    
    If DiaryForm.Cancelled Then Exit Sub
    
    m_Caller.CreateEvent EventType:=2006, EventText:=DiaryForm.Text, Params:=Params

Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "Diary Entry: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Private Sub DetermineBalanceEnabled()
    On Error GoTo Handler
    
    Me.MousePointer = vbDefault
    
    If IsHidden() Then
        cmdBalance.Enabled = False
    ElseIf m_Caller Is Nothing Then
        cmdBalance.Enabled = False
    ElseIf m_Caller.SelectedCustomer Is Nothing Then
        cmdBalance.Enabled = False
    ElseIf m_Caller.SelectedCustomer.Accounts.Count < 1 Then
        cmdBalance.Enabled = False
    Else
        Dim MenuControl As IMenuControl
        Set MenuControl = IDCContainer.IDCServiceManager.Retrieve(MENU_CONTROL)
        cmdBalance.Enabled = MenuControl.IsServiceEnabled(DISPLAY_SELECTED_BALANCE)
        'Me.MousePointer = IIf(cmdBalance.Enabled, vbDefault, vbHourglass)
'        Dim MenuFunctions As IFunctionObject
'        Set MenuFunctions = MenuControl
'        Dim Params As New Collection
'        Params.Add DISPLAY_SELECTED_BALANCE
'        Me.MousePointer = IIf(MenuFunctions.CallFunction("running", Params) = 0, vbDefault, vbHourglass)
        Me.MousePointer = IIf(MenuControl.MenuService(DISPLAY_SELECTED_BALANCE).RunningInstances = 0, vbDefault, vbHourglass)
    End If
    
    Dim OldEnabled As Boolean
    OldEnabled = fpList1.Enabled
    fpList1.Enabled = cmdBalance.Enabled
    
    ' Screen.ActiveForm always seems to be "Me", even if the MDI form has the focus. See the bodge
    ' in "BalanceReturned"
'    If Screen.ActiveForm Is Me And fpList1.Enabled And Not OldEnabled And fpList1.Visible Then
'        On Error Resume Next
'        Debug.Print "SCREEN", Screen.ActiveForm.Name
'        Debug.Print "MDI", IDCContainer.IDCMDIWindow.ActiveForm.Name
'        Debug.Print "BASE", ActiveForm.Name
'        fpList1.SetFocus
'    End If
Exit Sub

    Resume ' For debugging only

Handler:
    IDCContainer.IDCLog.Record "DetermineBalanceEnabled: " & Err.Description, idcSeverityError, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Public Sub DetermineDiaryEnabled()
    Dim CustObj As PIDCCustomerData.CCustomerData
    
    On Error Resume Next
    
    ' PILOT PROBLEM - for some reason SSTab1 was being disabled when CCustomerData completed
'    If Not SSTab1.Enabled Then
'        SSTab1.TabEnabled(2) = False
'    ElseIf m_Caller.PersonalCustomer Is Nothing Then
    If m_Caller Is Nothing Then
        SSTab1.TabEnabled(2) = False
    ElseIf m_Caller.PersonalCustomer Is Nothing Then
        SSTab1.TabEnabled(2) = False
    Else
'        SSTab1.TabEnabled(2) = (lstRelationshipManagers.ListCount > 0) Or (lstSegmentMarkers.ListCount > 0)
        SSTab1.TabEnabled(2) = True
    End If
    
    cmdDiaryEntry.Enabled = SSTab1.TabEnabled(2) And (lstRelationshipManagers.ListIndex >= 0)
End Sub

Private Sub SetupScreenHeader()
    
    On Error GoTo ErrHandler
 
    If IsHidden() Then Exit Sub
    
    Dim bcxData As GlobalBCX1: Set bcxData = IDCContainer.IDCServiceManager.Retrieve("BCXGlobalData")
       
    txtMembershipNo = m_Caller.MembershipNumber
        
    txtName = m_Caller.Name

    ' SIR 2664 - display the CALLER's CIS ID immediately after initialisation, as the advisers
    ' sometimes need to write it down if the caller is suspended. If the caller is linked
    ' to > 1 customers (e.g. Personal+Business), this may be changed after verification and
    ' customer selection is completed.
    
    txtCISID = m_Caller.CisId

    
    If m_Caller.Age > 0 Then
        txtAge = m_Caller.Age
    End If
    
    SetupTranche
    
    indCustomerIndicators.Enabled = True
    
Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "SetupScreenHeader: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Private Sub SetupTranche()
    Dim bcxData As GlobalBCX1: Set bcxData = IDCContainer.IDCServiceManager.Retrieve("BCXGlobalData")
    
    txtPhoneNumber = ""
    
    If Not m_Caller.Customer.ConsiderAsLive Then
        txtPhoneNumber = ""
    ElseIf m_Caller.BcallCustomer.PersonalCustomerDetails(1).Ind8 <> "" Then
        Dim unloadRow1 As CRecTrch
        For Each unloadRow1 In bcxData.gl_trch_array.AllRows
            If unloadRow1.trc_id = m_Caller.BcallCustomer.PersonalCustomerDetails(1).Ind8 Then
                txtPhoneNumber = Trim(unloadRow1.svc_dsc) & vbCrLf & Trim(unloadRow1.tlp_num)
                Exit For
            End If
        Next
    End If
    
    If m_Caller.BcallCustomer.PersonalCustomerDetails(1).Ind8 = "D" Then
        ' Realease 24, horrible bodge.
        ' The Business Direct managers are moving from Walsgrave to Gadbrook Park. For some time,
        ' there will be managers at both centres, and IDT must tell the advisers which centre
        ' the customer "belongs to".
        '
        ' Customers taken on after a certain date (TAKEON code in ref data) will be at the new site.
        ' Existing customers will migrate to Gadbrrok in alphabetical groups (MIGRATED code in refdata).
        
        Dim LRefData  As IIDCLocalRefData
        Dim LRefItems As IIDCLocalReferenceItems
        Dim Item As IIDCLocalReferenceItem
        Dim Loc1 As String
        Dim Loc2 As String
        Dim Migrated As String
        Dim TakeOn As Date
        Dim BusinessTakeOn As Date
        Dim Suffix As String
        Dim IsBusinessDirect As Boolean
        
        If m_Caller.BusinessCustomers.Count > 0 Then
            ' The IsBusinessDirect value is not used. However it is essetial to query
            ' m_Caller.BusinessCustomers(1).BusinessDirect as this forces the CPL record for
            ' the business to be read.
            
            IsBusinessDirect = m_Caller.BusinessCustomers(1).BusinessDirect
            
            Debug.Assert IsBusinessDirect
            
            BusinessTakeOn = m_Caller.BusinessCustomers(1).BcallCustomer.PersonalCustomerDetails(1).CusSgnTelDte
            
            Set LRefData = IDCContainer.IDCServiceManager.Retrieve(LOCAL_REF_DATA)
            Set LRefItems = LRefData.GetLocalRefItems("BDRMmove")
            
            
            For Each Item In LRefItems
                If Item.Code = "TAKEON" Then
                    TakeOn = CDate(Item.Description)
                ElseIf Item.Code = "MIGRATED" Then
                    Migrated = UCase(Item.Description)
                ElseIf Item.Code = "LOC1" Then
                    Loc1 = Item.Description
                ElseIf Item.Code = "LOC2" Then
                    Loc2 = Item.Description
                Else
                    IDCContainer.IDCLog.Record "Invalid entry in " & Item.Class & " local RefData: " & Item.Code, idcSeverityError, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT
                End If
            Next
            
            IDCContainer.IDCLog.Trace "Business Direct: TakeOn=" & TakeOn & ", Migrated=" & Migrated & ", Loc1=" & Loc1 & ", Loc2=" & Loc2
            
            If BusinessTakeOn >= TakeOn Then
                Suffix = Loc2
            ElseIf Migrated <> "" And UCase(Left(m_Caller.BusinessCustomers(1).Name, Len(Migrated))) <= Migrated Then    ' Defect 294
                Suffix = Loc2
            Else
                Suffix = Loc1
            End If
            
            txtPhoneNumber = Trim(unloadRow1.svc_dsc) & " " & Suffix & vbCrLf & Trim(unloadRow1.tlp_num)
        End If
    End If
End Sub

Public Sub IDVComplete()
    On Error GoTo ErrHandler
    
    If IsHidden() Then Exit Sub
    
    Dim Customer As IIDCCustomer
    
    ' populate the customer list
    
    cmbCustomers.Clear
    
    If m_Caller.BusinessCustomers.Count > 0 Then
    
        For Each Customer In m_Caller.Customers
            cmbCustomers.AddItem Customer.Name
        Next
        
        cmbCustomers.Visible = True
        cmbCustomers.TabStop = True
        frmCustomer.Caption = "C&ustomer"
        cmbCustomers.Text = m_Caller.SelectedCustomer.Name
    Else
        cmbCustomers.Visible = False
        cmbCustomers.TabStop = False
        frmCustomer.Caption = "Customer"
    End If
    
    DisplayCustomerData

    ' Now set up the remaining details on the "IVR tab"
    Dim Tel As IIDCTelephony
    Set Tel = IDCContainer.IDCServiceManager.Retrieve("Telephony")
    With Tel
        ' Could fail if there was something wrong with the account data
        On Error Resume Next
        If .FromAccountNumber <> "00000000" And .FromSortcode <> "" Then
            txtReference(0) = m_Caller.SelectedCustomer.Accounts(.FromSortcode & .FromAccountNumber).AccountShortName
        End If
        If .ToAccountNumber <> "00000000" And .ToSortcode <> "" Then
            txtReference(1) = m_Caller.SelectedCustomer.Accounts(.ToSortcode & .ToAccountNumber).AccountShortName
        End If
        On Error GoTo ErrHandler
    End With
    
    EnableAllControls

Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "Form.IDVComplete: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Public Sub DisplayCustomerData()
On Error GoTo ErrHandler
'Get the mainframe location and get the first Char
   
    If m_Caller Is Nothing Then Exit Sub
    If m_Caller.SelectedCustomer Is Nothing Then Exit Sub
    If IsHidden() Then Exit Sub
    
    If m_Caller.SelectedCustomer.PrimaryAccountSortCode <> "" Then
        txtPrimaryBranch = m_Caller.SelectedCustomer.PrimaryAccountSortCode & " " & _
            UCase(m_Caller.SelectedCustomer.PrimaryAccountBranchName)
        txtProcessingCentre = m_Caller.SelectedCustomer.ProcessingCentre
    End If
    
    txtCISID = m_Caller.SelectedCustomer.CisId
    
    ' Set the combobox to the correct list item (obviously will
    ' already be the right item if the Context Change resulted from
    ' the user clicking on the combo box).
    
    Dim ACust As IIDCCustomer
    Dim i     As Long
    If cmbCustomers.Visible Then
        i = 0
        For Each ACust In m_Caller.Customers
            If ACust Is m_Caller.SelectedCustomer Then
                ' This is the one. NB combo ListIndexes are zero based, collections one-based
                If cmbCustomers.ListIndex <> i And cmbCustomers.ListCount > i Then cmbCustomers.ListIndex = i
                Exit For
            End If
            i = i + 1
        Next
    End If
        
    txtBusinessIndicator.Visible = True
    
    If m_Caller.SelectedCustomer.Business Then
        If m_Caller.SelectedCustomer.BusinessDirect Then
            txtBusinessIndicator.Text = "BDIR"
        Else
            txtBusinessIndicator.Text = "BUSN"
        End If
        lblUpdatedLedger.Visible = True
        txtUpdatedLedger.Visible = True
        fpList1.ColFromName = "Balance"
        fpList1.ColWidth = 16
        fpList1.ColFromName = "Overdraft"
        fpList1.ColWidth = 15
        fpList1.ColFromName = "Ledger"
        fpList1.ColWidth = 16
        fpList1.ColHide = False
    Else
        txtBusinessIndicator.Text = "PERS"
        lblUpdatedLedger.Visible = False
        txtUpdatedLedger.Visible = False
        fpList1.ColFromName = "Balance"
        fpList1.ColWidth = 23.6
        fpList1.ColFromName = "Ledger"
        fpList1.ColWidth = 0
        fpList1.ColHide = True
        fpList1.ColFromName = "Overdraft"
        fpList1.ColWidth = 23.4
    End If
    
    DisplayReferIndicator
    
    'Centralisation
    'On the balance tab make sure the summary screen is displayed rather than any of the detail screens
    frmLoanBalance.Visible = False
    frmLoanBalance.Enabled = False
    frmDetailBalance.Visible = False
    frmDetailBalance.Enabled = False
    frmSummaryBalance.Visible = True
    frmSummaryBalance.Enabled = True
    'also clear the values on summary screen (as they are still populated if the affiliated customer has no accounts)
    txtBalance.Text = ""
    txtOverdraft.Text = ""
    txtUpdatedLedger.Text = ""
    
    DisplayAccounts
    
    ' SMR 6/7/98 - error condition - should never happen
    ' but just in case...
    DetermineBalanceEnabled
Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record _
        "DisplayCustomerData: " & Err.Description & vbCrLf & _
        "             CIS ID: " & m_Caller.CisId & vbCrLf & _
        "   SelectedCustomer: " & m_Caller.SelectedCustomer.CisId, _
            idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Public Sub DisplayReferIndicator()
    ' Search for a Refer Stream.
    If Not m_Caller.SelectedCustomer.Customer.EligibleAccounts Is Nothing Then
        With m_Caller.SelectedCustomer.Customer.EligibleAccounts
            If .Count > 0 Then
                If Trim(.Item(1).ReferStream) <> "" Then
                    txtBusinessIndicator = Trim(.Item(1).ReferStream)
                End If
            End If
        End With
    End If
End Sub

Public Sub Clear()
    txtName = ""
    txtCISID = ""
    txtAge = ""
    ClearCustomerIndicators
    txtPrimaryBranch = ""
    txtOverdraft = ""
    txtBalance = ""
    txtUpdatedLedger = ""
    txtMembershipNo = ""
    txtBusinessIndicator = ""
    txtProcessingCentre = ""
    txtPhoneNumber = ""
    cmbCustomers.Clear

    fpList1.Clear
    txtIVRReasonForTransfer = ""
    txtLastIVRServiceUsed = ""
    chkIdentifiedByIVR.Value = vbUnchecked
    chkIdentifiedByIVR.Caption = "Identified By IVR"
    chkVerifiedByIVR.Value = vbUnchecked
    txtTransAccount(0) = ""
    txtTransAccount(1) = ""
    txtTransSortCode(0) = ""
    txtTransSortCode(1) = ""
    txtReference(0) = ""
    txtReference(1) = ""
    txtTransAmount = ""
    txtTransDate = ""
    txtBarclaycard = ""
    
    'Centralisation - clear contents of the Loan & non-loan frames
    txtBalance1.Text = ""
    txtBalance2.Text = ""
    txtOverdraft1.Text = ""
    txtOverdraft2.Text = ""
    txtStatement1.Text = ""
    txtStatement2.Text = ""
    txtLoanBalance.Text = ""
    txtLoanInterest.Text = ""
    txtLoanTotal.Text = ""
    txtLoanSubtotal.Text = ""
    txtLoanEarlySett.Text = ""
    
    'December 2000
    lblWarning.Visible = False
    lblWarningText.Visible = False
    
    lstSegmentMarkers.Clear
    lstRelationshipManagers.Clear
    
    Set m_Caller = Nothing
    
    DisableAllControls
End Sub

Public Sub DisableAllControls()
    cmdBalance.Enabled = False
    If Not m_Caller Is Nothing Then m_Caller.Parent.ClearEnabled = False
    SSTab1.Tab = 0
    SSTab1.Enabled = False
    DetermineDiaryEnabled
    fpList1.Enabled = False
    cmbCustomers.Enabled = False
    txtName.Enabled = False
    txtMembershipNo.Enabled = False
    txtCISID.Enabled = False
    txtAge.Enabled = False
    txtBusinessIndicator.Enabled = False
    txtPrimaryBranch.Enabled = False
    txtPhoneNumber.Enabled = False
    txtProcessingCentre.Enabled = False
    
    'Centralisation - disable and make invisible the Loan & non-loan frames and make visible the summary frame
    frmLoanBalance.Visible = False
    frmLoanBalance.Enabled = False
    frmDetailBalance.Visible = False
    frmDetailBalance.Enabled = False
    frmSummaryBalance.Visible = True
    frmSummaryBalance.Enabled = False
End Sub

Public Sub EnableAllControls()
    
    If IsHidden() Then Exit Sub
    
    ' ADE 00035276: If SSTab1 is not enabled, then cmdBalance cannot be enabled.
    SSTab1.Enabled = True
    
    DetermineBalanceEnabled
    DetermineDiaryEnabled
    
    cmbCustomers.Enabled = cmdClear.Enabled
    txtName.Enabled = True
    txtMembershipNo.Enabled = True
    txtCISID.Enabled = True
    txtAge.Enabled = True
    txtBusinessIndicator.Enabled = True
    txtPrimaryBranch.Enabled = True
    txtPhoneNumber.Enabled = True
    txtProcessingCentre.Enabled = True
    'Centralisation
    frmSummaryBalance.Visible = True
    frmSummaryBalance.Enabled = True
End Sub


Public Property Let ClearEnabled(ByVal X As Boolean)
    If X Then
        If IsHidden() Then Exit Property
        
        cmdClear.Enabled = True
        cmbCustomers.Enabled = True
        
        DetermineBalanceEnabled
    Else
        cmdClear.Enabled = False
        cmbCustomers.Enabled = False
        DetermineBalanceEnabled
    End If
End Property
Public Property Get ClearEnabled() As Boolean
    ClearEnabled = cmdClear.Enabled
End Property


Public Sub DisplayAccounts()
On Error GoTo ErrHandler
    fpList1.Clear
    
    If IsHidden() Then Exit Sub
    
    Dim tabStr As String: tabStr = Chr(9)
    Dim Index As Integer: Index = 0
    Dim Account As IIDCAccount
    
'    For Each Account In m_Caller.SelectedCustomer.GetAccounts("AE") 'Centralisation
'    For Each Account In m_Caller.SelectedCustomer.GetAccounts("BE") 'Centralisation
    For Each Account In m_Caller.SelectedCustomer.GetAccounts("BAL") ' OpenPlan
        Dim listStr As String
        listStr = Account.AccountType & tabStr & Account.AccountShortName(acnComposite) & tabStr & _
            Account.Sortcode & tabStr & Account.AccountNumber
        fpList1.AddItem listStr, Index
        Index = Index + 1
    Next
    
    ' select the primary account (if any)
    ' SIR2942 - do not select if no accounts
    
    If m_Caller.SelectedCustomer.Accounts.Count > 0 Then
        With fpList1
            .Row = 0
            .Selected = True
        End With
    End If
    
    DetermineBalanceEnabled
    
    IDCContainer.IDCLog.Trace "CustomerContext - Account Table Displayed"
Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "DisplayAccounts: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Public Sub BalanceReturned()
    
    On Error GoTo ErrHandler
    If IsHidden() Then Exit Sub
    
    Dim MenuControl As IMenuControl
    Set MenuControl = IDCContainer.IDCServiceManager.Retrieve(MENU_CONTROL)

    If Not MenuControl.GetAuthorisation("DispSelBal", DoNotVerify:=True) Then Exit Sub
    
    With fpList1
        .ColFromName = "Balance"
        .ColList = m_Caller.SelectedAccount.CurrentClearedForFundsBalance
        txtBalance = m_Caller.SelectedAccount.CurrentClearedForFundsBalance
        
        .ColFromName = "Overdraft"
        If m_Caller.SelectedAccount.OverdraftLimit <> "0" Then
            .ColList = m_Caller.SelectedAccount.OverdraftLimit
            txtOverdraft = m_Caller.SelectedAccount.OverdraftLimit
        End If
        
        .ColFromName = "ledger"
        .ColList = m_Caller.SelectedAccount.UpdatedLedgerBalance
        txtUpdatedLedger = m_Caller.SelectedAccount.UpdatedLedgerBalance
        
        .ColFromName = "Balance"
        If gRight(.ColList, 1) <> "R" And .ColList <> "UNAVAILABLE" Then   ' 4Q96
            .ColList = .ColList & "     " ' 5 spaces is ~ the width of 'DR'
        End If
        .ColFromName = "Ledger"
        If gRight(.ColList, 1) <> "R" And .ColList <> "UNAVAILABLE" Then   ' 4Q96
            .ColList = .ColList & "     " ' 5 spaces is ~ the width of 'DR'
        End If
        
        .ColFromName = "IsBaProduct"
        .ColList = CStr(m_Caller.SelectedAccount.IsBAProduct)
        
        ' Bodge Alert!!    Bodge Alert!!
        ' The handling of the enabling/disabling the balance button/account list has been rationalised in 2Q2000.
        ' The button/list is now disabled when the "DisplaySelectedBalance" service starts, end re-enabled when
        ' the service ends.
        ' However when we are at this point in the code, DisplaySelectedBalance is still running (this subroutine
        ' is called indirectly by DisplaySelectedBalance). We want to set the focus back to the account list, but
        ' we can't because we have not yet been enabled.
        '
        ' Could let DetermineBalanceEnabled set the focus to the account list when it makes the account list enabled, but
        ' this does not work - it will give the focus to CustomerContext even if we do not want it to. Enable & SetFocus it
        ' here instead, although it isn't very satisfactory.
        '
        .Enabled = True
        If .Enabled And .Visible Then .SetFocus
    End With
        
    ' OpenPlan - display balances from the Miscellaneous Accounts
    If m_Caller.SelectedAccount.IsBAProduct Then
        
        ' Normal BA product
        
        'December 2000 Display uncleared cheques warning message if needed.
        Dim curUnclearedAmount As Currency
        Dim strUnclearedAmount As String
        Dim curClearedAmount As Currency
        Dim strClearedBalance As String
        With m_Caller.SelectedAccount.AccountData
            If Not .AccountInfo(gaeCombUnclFATELadder) Is Nothing And Not .AccountInfo(gaeAuthLimits) Is Nothing Then
                If .AccountInfo(gaeCombUnclFATELadder).Count > 0 And .AccountInfo(gaeAuthLimits).Count > 0 Then
                    If .AccountInfo(gaeCombUnclFATELadder).Item(1).AmountSchedules.Item(2).Amount < Abs(.AccountInfo(gaeAuthLimits).Item(1).Limit(gaeAcUnclearedLim)) Then 'ADE 00039997
                        curUnclearedAmount = .AccountInfo(gaeCombUnclFATELadder).Item(1).AmountSchedules.Item(2).Amount
                    Else
                        curUnclearedAmount = .AccountInfo(gaeAuthLimits).Item(1).Limit(gaeAcUnclearedLim)
                    End If
                    If curUnclearedAmount <> 0 Then
                        If IsNumeric(m_Caller.SelectedAccount.CurrentClearedForFundsBalance) Then
                            curClearedAmount = CCur(m_Caller.SelectedAccount.CurrentClearedForFundsBalance) - Abs(curUnclearedAmount)
                            strClearedBalance = FormatDebit(curClearedAmount, True)
                            strUnclearedAmount = FormatDebit(curUnclearedAmount, False)
                            lblWarning.Visible = True
                            lblWarningText.Caption = FormatWarningText(strUnclearedAmount, strClearedBalance)
                            lblWarningText.Visible = True
                            fpList1.ColFromName = "UnclearedAmount"
                            fpList1.ColList = strUnclearedAmount
                            fpList1.ColFromName = "ClearedBalance"
                            fpList1.ColList = strClearedBalance
                        End If
                    End If
                End If
            End If
        End With
    End If
    
    ' Now simulate a SelChange operation so that the "warnings" are refreshed where necessary.
    fpList1_SelChange fpList1.ListIndex
    
Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "Balance Returned: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Private Function FormatWarningText(strUnclearedAmount As String, strClearedBalance As String) As String
    Dim strWarningText As String
    strWarningText = "Balance includes £" & strUnclearedAmount & " of uncleared cheques" & vbCrLf
    strWarningText = strWarningText & "Cleared Balance is £" & strClearedBalance
    FormatWarningText = strWarningText
End Function

Private Sub RelationshipInformationReturned(ByVal Customer As IIDCCustomer)
    On Error GoTo ErrHandler
    
    Dim CustObj As PIDCCustomerData.CCustomerData
    Set CustObj = Customer.Customer
    
    If m_Caller Is Nothing Then Exit Sub
    If IsHidden() Then Exit Sub
    
    Dim BusUnit As CBusinessUnit
    
    lstRelationshipManagers.Clear
    For Each BusUnit In CustObj.BusinessUnits
        With BusUnit
            If .IsCurrent And .state <> Deleted And .CustomerBusinessUnitRelationTypeCode = "R" Then
                lstRelationshipManagers.AddItem _
                        Trim(.EmployeeName) & " " & Trim(.EmployeeIts) & " (" & Trim(.BusinessUnitName) & ")" & vbTab & _
                        Trim(.BusinessUnitId) & vbTab & _
                        Trim(.EmployeeNumber) & vbTab & _
                        Trim(.EmployeeName) & vbTab & _
                        Trim(.EmployeeIts)
            End If
        End With
    Next
    
    If lstRelationshipManagers.ListCount = 1 Then
        lstRelationshipManagers.ListIndex = 0
    End If
    
'    Set BusUnit = FindBusinessUnit(Customer.Customer)
'
'    If Not BusUnit Is Nothing Then
'        With BusUnit
'            ' NB: Form design liable to change, so not all of these text boxes are visible.
'            txtBUEmpName = Trim(.EmployeeIts) & " " & Trim(.EmployeeName)
'            txtBUEmpNumber = Trim(.EmployeeNumber)
'            txtBUName = Trim(.BusinessUnitName)
'        End With
'    End If
    
Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "RelationshipInformationReturned: " & Err.Description & vbCrLf & CustObj.Dump(), idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Public Sub CustomerDataRetrieved(ByVal Customer As IIDCCustomer)
On Error GoTo Handler

    Dim CustCon As IIDCCUSTOMERCONTEXTLib.IDCCustomerContext
    Set CustCon = GetCustCon()
    
    Dim i As Integer
    Dim Sortcode As String
    Dim AccountNumber As String
    Dim ShortName As String
    Dim NewShortName As String
    Dim Account As IIDCAccount
    
    ' The customer is question could be the selected customer, or the personal
    ' customer (and the selected and personal customers may be the same thing).
    
    If Customer Is CustCon.Caller.PersonalCustomer Then
        If Customer.Customer.errorObject(ErrObj_CUSTOMER) Is Nothing Then
            If Customer Is m_Caller.PersonalCustomer Then
                RelationshipInformationReturned Customer
            End If
        Else
            IDCContainer.IDCLog.Record "CustomerDataRetreieved: " & Customer.Customer.errorObject.Dump, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
        End If
        
        SetCustomerIndicators
        
        DisplayReferIndicator
    End If
    
    If Customer Is CustCon.Caller.SelectedCustomer Then
        ' Novenber 2001: Check if any account short names have changed.
        
        If Not IsHidden() Then
            For i = 0 To fpList1.ListCount - 1
                fpList1.Row = i
                
                fpList1.ColFromName = "SortCode"
                Sortcode = fpList1.ColList
                fpList1.ColFromName = "AccountNumber"
                AccountNumber = fpList1.ColList
                fpList1.ColFromName = "cus_acc_sht_nme"
                ShortName = fpList1.ColList
                
                Debug.Print i & " " & Sortcode & " " & AccountNumber & " " & ShortName & "."
                
                Set Account = Customer.FindAccount(Sortcode, AccountNumber)
                If Account Is Nothing Then
                    'IDCContainer.IDCLog.Trace "Could not find " & SortCode & " " & AccountNumber & " for " & Customer.cisID
                Else
                    NewShortName = Trim(Account.AccountShortName(acnComposite))
                    
                    If NewShortName <> Trim(ShortName) Then
                        IDCContainer.IDCLog.Trace "Account " & Sortcode & " " & AccountNumber & " for " & Customer.CisId & " changed name from """ & ShortName & """ to """ & NewShortName & """"
                        fpList1.ColFromName = "cus_acc_sht_nme"
                        fpList1.ColList = NewShortName
                    End If
                End If
            Next
        End If
    End If
Exit Sub

Handler:
    Dim ErrNum As Long: ErrNum = Err.Number
    Dim ErrDesc As String: ErrDesc = Err.Description
    IDCContainer.IDCLog.Record "CustomerDataRetrieved: " & ErrDesc, idcSeverityError, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT
End Sub

Public Sub BondSummaryDetailsRetrieved(ByVal Customer As IDCCustomer)
    ' We now know the sub-product code. Redisplay account list to include these details.
    If Customer Is m_Caller.SelectedCustomer Then DisplayAccounts
End Sub


Public Sub IVRInitialise(TELEPHONY As IIDCTelephony)
    On Error GoTo ErrHandler
    If IsHidden() Then Exit Sub
    
    With TELEPHONY
        If Trim(.CustomerName) <> "" Then
           txtName = Prettify(.CustomerName)
        End If
        txtIVRReasonForTransfer = .IVRReason
        txtLastIVRServiceUsed = .LastIVRTransactionType
        
        lblBarclaycard.Visible = False
        txtBarclaycard.Visible = False
        lblAccounts(0).Visible = False
        lblAccounts(1).Visible = False
        lblAmount.Visible = False
        lblEffectiveDate.Visible = False
        txtTransAccount(0).Visible = False
        txtTransAccount(1).Visible = False
        txtTransSortCode(0).Visible = False
        txtTransSortCode(1).Visible = False
        txtReference(0).Visible = False
        txtReference(1).Visible = False
        txtTransAmount.Visible = False
        txtTransDate.Visible = False
        
        If .FromAccountNumber <> "00000000" And .FromSortcode <> "" Then
               lblAccounts(0).Visible = True
               txtTransAccount(0).Visible = True
               txtTransSortCode(0).Visible = True
               txtReference(0).Visible = True
               txtTransAccount(0) = .FromAccountNumber
               txtTransSortCode(0) = .FromSortcode
               txtTransAmount.Visible = True
               lblAmount.Visible = True
        End If
        
        If .ToAccountNumber <> "00000000" And .ToSortcode <> "" Then
               lblAccounts(1).Visible = True
               txtTransAccount(1).Visible = True
               txtTransSortCode(1).Visible = True
               txtReference(1).Visible = True
               txtTransAccount(1) = .ToAccountNumber
               txtTransSortCode(1) = .ToSortcode
        End If
        
        If .LastIVRTransactionCode = tcBarclaycardBillPayment Then
            lblBarclaycard.Visible = True
            txtBarclaycard.Visible = True
            txtTransDate.Visible = True
            lblEffectiveDate.Visible = True
        End If
        
        txtBarclaycard = .BarclaycardNumber
        If .TransferAmount <> 0 Then
            txtTransAmount = Format(.TransferAmount, "0.00")
        End If
        
        ' SIR 3711: IsDate("31/09") returns True! Changed to check variant's type.
        ' (The Telephony object will set it to a date if it was in the correct format,
        ' otherwise a string).
        If VarType(.EffectiveTransferDate) = vbDate Then
            txtTransDate = Format(.EffectiveTransferDate, "dd/mm/yyyy")
        Else
            txtTransDate = .EffectiveTransferDate
        End If
        
        If .CustomerIdentified Then
            chkIdentifiedByIVR.Value = vbChecked
            If .CardNumber = "" Then
                chkIdentifiedByIVR.Caption = "Identified By IVR, by Membership Number"
            Else
                chkIdentifiedByIVR.Caption = "Identified By IVR, by Bank Card Number"
            End If
        Else
            chkIdentifiedByIVR.Value = vbUnchecked
        End If
        
        If .CustomerVerified Then
            chkVerifiedByIVR.Value = vbChecked
        Else
            chkVerifiedByIVR.Value = vbUnchecked
        End If
        
        Dim IVRCall As Boolean
        IVRCall = .AttachedData <> ""
        If Not IVRCall Then
            SSTab1.Tab = 0
        Else
            SSTab1.Tab = 1
        End If
    End With
Exit Sub

ErrHandler:
    IDCContainer.IDCLog.Record "IVRInitialise: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
End Sub

Public Sub LockCountChanged(ByVal NewValue As Long)
    On Error Resume Next
    If Not m_Caller Is Nothing Then m_Caller.Parent.ClearEnabled = (NewValue = 0)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    ' Bodge Alert!!    Bodge Alert!!
    ' For some reason, the sizes of the FarPoint ListBoxes
    ' seem to be overwritten. Set it here.
    With lstRelationshipManagers
        .Height = 255
        .Width = 4380
    End With
    With lstSegmentMarkers
        .Height = 840
        .Width = 5250
    End With
End Sub

Private Sub txtBalance_Change()
    'Centralisation
On Error GoTo Handler
    If Trim(txtBalance.Text) <> "" And Trim(txtBalance.Text) <> cError Then
        If m_Caller.SelectedAccount.ProductReference.InGroup(idcACIBond) Then
            cmdDetail.Enabled = False       ' R26
        Else
            cmdDetail.Enabled = True
        End If
    Else
        cmdDetail.Enabled = False
    End If
Exit Sub
Handler:
      IDCContainer.IDCLog.Record "Error in txtBalance_Change: " & Err.Description, idcSeverityError, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT
End Sub

Private Sub DisplayDetail()
    frmSummaryBalance.Visible = False
'    frmSummaryBalance.Enabled = False 'leave it enabled so that the fplist and balance button remain enabled

    If Not m_Caller.SelectedAccount.IsBAProduct Then
        DisplayMiscAccountDetail
    ElseIf m_Caller.SelectedAccount.ProductReference.ClassCode = "LN" Then
        DisplayLoanDetail
    Else
        DisplayNonLoanDetail
    End If
End Sub

Private Sub DisplayMiscAccountDetail()
On Error GoTo ErrHandler
    With m_Caller.SelectedAccount 'ADE 37152
        txtMiscDetail = "Details for " & .Sortcode & " " & .AccountNumber & " " & .AccountShortName
    End With
            
    frmMiscBalance.Visible = True
    frmMiscBalance.Enabled = True
    
    txtMiscAvailable.Text = ""
    txtMiscBalance.Text = ""
    txtMiscInitial.Text = ""
    
    Dim CustCon As IIDCCUSTOMERCONTEXTLib.IDCCustomerContext
    Set CustCon = GetCustCon()
    
    Dim Cust As IIDCCustomer
    Dim CustData As PIDCCustomerData.CCustomerData
    Dim MiscBal As PIDCCustomerData.CMiscBalance
    
    For Each Cust In CustCon.Caller.Customers
        If Not Cust.FindAccount(m_Caller.SelectedAccount.Sortcode, m_Caller.SelectedAccount.AccountNumber) Is Nothing Then
            Set CustData = Cust.Customer
            
            For Each MiscBal In CustData.MiscBalances
                Dim Sortcode As String
                Dim AccNumber As String
                Sortcode = Left(MiscBal.GenericAccountNumber, 6)
                AccNumber = Mid(MiscBal.GenericAccountNumber, 7, 8)
                If Sortcode = m_Caller.SelectedAccount.Sortcode And AccNumber = m_Caller.SelectedAccount.AccountNumber Then
                    txtMiscAvailable.Text = FormatDebitOrBlank(MiscBal.AmountAvailable, False)
                    txtMiscBalance.Text = FormatDebitOrBlank(MiscBal.Balance, False)
                    txtMiscInitial.Text = FormatDebitOrBlank(MiscBal.InitialAmount, False)
                    Exit Sub
                End If
            Next
        End If
    Next
    
    ' Account not found
    
    txtMiscAvailable.Text = cError
    txtMiscBalance.Text = cError
    txtMiscInitial.Text = cError

Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "DisplayMiscDetail: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
    'redisplay previous details
    frmMiscBalance.Visible = False
    frmMiscBalance.Enabled = False
    frmSummaryBalance.Visible = True
    frmSummaryBalance.Enabled = True
End Sub

    
Private Sub DisplayLoanDetail()
On Error GoTo ErrHandler
    Dim curTotal As Currency: curTotal = 0
    
    With m_Caller.SelectedAccount 'ADE 37152
        txtLoanDetail.Text = "Details for loan " & .Sortcode & " " & .AccountNumber & " " & .AccountShortName
    End With
            
    Dim LocalRefData As IIDCLocalRefData
    Dim LocalRefItems As IIDCLocalReferenceItems
    Set LocalRefData = IDCContainer.IDCServiceManager.Retrieve("LocalRefData")
    If LocalRefData Is Nothing Then
        Err.Raise vbObjectError + 1, "Customer Context ServiceForm: DisplayLoanDetail", "Could not retrieve local reference data"
    End If
    Set LocalRefItems = LocalRefData.GetLocalRefItems("DoNotUse", , "LoanDetails", True)
    Dim flgUseOldTransaction As Boolean: flgUseOldTransaction = False
    If LocalRefItems.Count <> 0 Then
        If LocalRefItems.Item(1).IsCurrent Then
            'Use old transaction
            flgUseOldTransaction = True
        End If
    End If

    With m_Caller.SelectedAccount.AccountData
        frmLoanBalance.Visible = True
        frmLoanBalance.Enabled = True
        'Clear all fields
        txtLoanBalance.Text = ""
        txtLoanInterest.Text = ""
        txtLoanTotal.Text = ""
        txtLoanSubtotal.Text = ""
        txtLoanEarlySett.Text = ""
        
        If Not flgUseOldTransaction Then
            IDCContainer.IDCLog.Trace "DisplayLoanDetail: EarlySettlement DAL.OESDetails:" & vbCrLf & .AccountInfo(gaeLoanDetails).Dump()
        End If
        
        If Not .errorObject Is Nothing Then 'Overall error
            txtLoanBalance.Text = cError
            txtLoanInterest.Text = cError
            txtLoanTotal.Text = cError
            txtLoanSubtotal.Text = cError
            txtLoanEarlySett.Text = cError
        Else
            If flgUseOldTransaction Then
                
               'Outstanding Balance
                If Not .AccountInfo(gaeBalances) Is Nothing Then
                    If .AccountInfo(gaeBalances).Count > 0 Then
                        curTotal = .AccountInfo(gaeBalances).Item(1).Balance(gaeRunningLedger)
                        txtLoanBalance.Text = FormatDebit(curTotal, True)
                    Else
                        txtLoanBalance.Text = AccountError(.ResponseCode(gaeBalances))
                    End If
                End If
                'Outstanding Interest
                If Not .AccountInfo(gaeDRIntPeriodic) Is Nothing Then
                    If .AccountInfo(gaeDRIntPeriodic).Count > 0 Then
                        txtLoanInterest.Text = FormatDebit(.AccountInfo(gaeDRIntPeriodic).Item(1).GrossInterest, True)
                        'long total carried forward from above calculation
                        curTotal = curTotal + .AccountInfo(gaeDRIntPeriodic).Item(1).GrossInterest
                    Else
                        txtLoanInterest.Text = AccountError(.ResponseCode(gaeDRIntPeriodic))
                    End If
                End If
                
                txtLoanSubtotal = FormatDebit(curTotal, True)
                
                txtLoanEarlySett.Text = "0.00"
                    
                If txtLoanBalance.Text <> cError And txtLoanInterest.Text <> cError Then
                    txtLoanTotal.Text = FormatDebit(curTotal, True)
                Else
                    txtLoanTotal.Text = cError
                End If
            Else
                With .AccountInfo(gaeLoanDetails)
                    'Outstanding Balance
                    curTotal = .LedgerBalance
                    txtLoanBalance.Text = FormatDebit(.LedgerBalance, True)
                    
                    'Outstanding Interest
                    curTotal = curTotal + .DebitInterestDue
                    txtLoanInterest.Text = FormatDebit(.DebitInterestDue, True)
                    
                    txtLoanSubtotal = FormatDebit(curTotal, True)
                    
                    'Early Settlement Interest
                    curTotal = curTotal + .EarlySettlementInterestToday
                    txtLoanEarlySett.Text = FormatDebit(.EarlySettlementInterestToday, True)
                    
                    txtLoanTotal = FormatDebit(curTotal, True)
                End With
            End If
        End If
    End With
Exit Sub
ErrHandler:
    IDCContainer.IDCLog.Record "DisplayLoanDetail: " & Err.Description, idcSeverityWarning, idcClassBusinessApplet, EV_CUSTOMER_CONTEXT_ERR_IN, 0
    'redisplay previous details
    frmDetailBalance.Visible = False
    frmDetailBalance.Enabled = False
    frmSummaryBalance.Visible = True
    frmSummaryBalance.Enabled = True
End Sub

Private Sub DisplayNonLoanDetail()
    Dim curTotal As Currency: curTotal = 0
    With m_Caller.SelectedAccount 'ADE 37152
        txtDetail.Text = "Details for " & .Sortcode & " " & .AccountNumber & " " & .AccountShortName
    End With
    
    With m_Caller.SelectedAccount.AccountData
        frmDetailBalance.Visible = True
        frmDetailBalance.Enabled = True
        'clear all fields
        txtBalance1.Text = ""
        txtBalance2.Text = "0.00"
        txtOverdraft1.Text = "0"
        txtOverdraft2.Text = "0.00"
        txtStatement1.Text = ""
        txtStatement2.Text = "0.00"
        
        If OverAllAccountError(.ResponseCode) Then 'Overall error
            'Should not happen
            txtBalance1.Text = cError
            txtBalance2.Text = cError
            txtOverdraft1.Text = cError
            txtOverdraft2.Text = cError
            txtStatement1.Text = cError
            txtStatement2.Text = cError
        Else
        
            'Outstanding Balance
            If Not .AccountInfo(gaeBalances) Is Nothing Then
                If .AccountInfo(gaeBalances).Count > 0 Then
                    curTotal = .AccountInfo(gaeBalances).Item(1).Balance(gaeRunningClrdFate)
                    txtBalance1.Text = FormatDebit(curTotal, True)
                Else
                    txtBalance1.Text = AccountError(.ResponseCode(gaeBalances))
                End If
            End If
            
            'Receipts not yet available for withdrawal
            If Not .AccountInfo(gaeCombUnclFATELadder) Is Nothing And _
               Not .AccountInfo(gaeAuthLimits) Is Nothing Then
                If .AccountInfo(gaeCombUnclFATELadder).Count > 0 And _
                   .AccountInfo(gaeAuthLimits).Count > 0 Then
                    curTotal = .AccountInfo(gaeCombUnclFATELadder).Item(1).AmountSchedules.Item(2).Amount
                    'ADE 00039306
                    Dim AuthLimitAmount As Currency
                    AuthLimitAmount = .AccountInfo(gaeAuthLimits).Item(1).Limit(gaeAcUnclearedLim)
                    If curTotal < Abs(AuthLimitAmount) Then
                        curTotal = 0
                    Else
                        curTotal = curTotal + AuthLimitAmount
                    End If
'                        curTotal = curTotal + .AccountInfo(gaeAuthLimits).Item(1).Limit(gaeAcUnclearedLim) 'ADE 00039239
                    txtBalance2.Text = FormatDebit(curTotal, False)
                Else
                    If AccountError(.ResponseCode(gaeCombUnclFATELadder)) = cError Or _
                       AccountError(.ResponseCode(gaeAuthLimits)) = cError Then
                        txtBalance2.Text = cError
                    Else
                        txtBalance2.Text = FormatDebit(0#, False)
                    End If
                End If
            End If
            
            'Overdraft Limit
            If Not .AccountInfo(gaeAuthLimits) Is Nothing Then
                If .AccountInfo(gaeAuthLimits).Count > 0 Then
                    curTotal = .AccountInfo(gaeAuthLimits).Item(1).Limit(gaeOverdraftLim)
                    'txtOverdraft1.Text = FormatOverdraft(curTotal)
                    'ADE #00039242 - JK
                    txtOverdraft1.Text = FormatDebit(curTotal, True)
                Else
                    txtOverdraft1.Text = Format(AccountError(.ResponseCode(gaeAuthLimits)), "0")
                End If
            End If
            
            'Using Overdraft ... withdraw up to
            If txtOverdraft1.Text <> cError Then
                If Not .AccountInfo(gaeBalances) Is Nothing Then
                    If .AccountInfo(gaeBalances).Count > 0 Then
                        curTotal = .AccountInfo(gaeBalances).Item(1).Balance(gaeFundsAvailable)
                        txtOverdraft2.Text = FormatDebit(curTotal, False)
                    Else
                        txtOverdraft2.Text = AccountError(.ResponseCode(gaeBalances))
                    End If
                End If
            End If
            
            'Updated Statement Balance
            If Not .AccountInfo(gaeBalances) Is Nothing Then
                If .AccountInfo(gaeBalances).Count > 0 Then
                    curTotal = .AccountInfo(gaeBalances).Item(1).Balance(gaeRunningLedger)
                    txtStatement1.Text = FormatDebit(curTotal, True)
                Else
                    txtStatement1.Text = AccountError(.ResponseCode(gaeBalances))
                End If
            End If
            
            'Outstanding Transactions
            If Not .AccountInfo(gaeOutstAuths) Is Nothing Then
                If .AccountInfo(gaeOutstAuths).Count > 0 Then
                    Dim i As Integer
                    curTotal = 0
                    For i = 1 To .AccountInfo(gaeOutstAuths).Count
                        If .AccountInfo(gaeOutstAuths).Item(i).AuthCommitted Then
                            curTotal = curTotal + .AccountInfo(gaeOutstAuths).Item(i).TotalAmount
                        End If
                    Next i
                    txtStatement2.Text = FormatDebit(curTotal, False)
                Else
                    txtStatement2.Text = AccountError(.ResponseCode(gaeOutstAuths))
                End If
            End If
        End If
    End With
End Sub

Private Function FormatDebitOrBlank(ByVal Amount As Variant, ByVal AddDR As Boolean) As String
    If Amount = "" Then Exit Function
    FormatDebitOrBlank = FormatDebit(Amount, AddDR)
End Function

Private Function FormatDebit(ByVal Amount As Currency, ByVal AddDR As Boolean) As String
    Dim strAmount As String
    Dim curAmount As Currency
    curAmount = Abs(Amount) 'Remove minus sign
    strAmount = Format(curAmount, "0.00")
    
    If AddDR Then
        If Amount < 0 Then
            strAmount = strAmount & "DR" 'Add Debit indicator
        End If
    End If
    FormatDebit = strAmount
End Function

Private Function FormatOverdraft(Amount As Currency) As String
    FormatOverdraft = Format(Abs(Amount), "0") 'remove sign and decimal places
End Function

Private Function OverAllAccountError(ErrorCode As Long) As Boolean
    If ErrorCode <> 111000 And _
       ErrorCode <> 111002 And _
       ErrorCode <> 111005 And _
       Not (ErrorCode >= 119000 And ErrorCode <= 119999) Then
        'overall response code = failure
        'I've included 111008 as a failure as this indicates no data has been returned - for a balance enquiry this must be a failure
        OverAllAccountError = True
    Else
        OverAllAccountError = False
    End If
End Function

Private Function AccountError(ErrorCode As Long) As String
    If ErrorCode <> 111000 And Not (ErrorCode >= 119000 And ErrorCode <= 119999) Then  'Unsuccessful
        'overall response code = failure
        'I've included 111008 as a failure as this indicates no data has been returned - for a balance enquiry this must be a failure
        AccountError = cError
    Else
        AccountError = FormatDebit(0, False)
    End If
End Function




Public Sub X()
	MsgBox str & Thing.Func(a, b)
	MsgBox str + Thing.Func(a, b)
	MsgBox str & Thing.Func
	MsgBox str & Thing.Func()
	MsgBox Func(a,b)
	MsgBox Thing.Func(a,b)
	MsgBox Thing.Func(a,b)(c,d).X(e)
	MsgBox Str1 & Err.Desc & Str2
	MsgBox Str1 & Err.Desc & " " & Erl & Str2
	If Not Thing.Status = Value Then MsgBox xx
	If Not Not + Not Thing.Status = - Not - Value Then MsgBox xx
	If Not Not Not Thing.Status = Not Value Then MsgBox xx
	MsgBox - a ^ b
	MsgBox - a * b
	MsgBox Not a And B
	MsgBox Not a = B
	dim X as Thing
	Dim Y as New Thing
	ReDim x(12) as New XXXXX
	set x = new y
	ReDim matSubComponentes(lIndice).dValorMensal(NumMes)
	ReDim This
End Sub

