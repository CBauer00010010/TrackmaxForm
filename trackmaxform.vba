' TrackmaxForm is a program to assist users in creating a valid request form
' Copyright (C) 2018 Christopher Ryan Bauer
'
' This program is free software: you can redistribute it and/or modify it under
' the terms of the GNU General Public License as published by the Free Software
' Foundation, either version 3 of the License, or (at your option) any later
' version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT
' ANY WARRANTY; without even the implied warranty of  MERCHANTABILITY or FITNESS
' FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License along with
' this program.  If not, see <http://www.gnu.org/licenses/>.


' ------------------------------------------------------------------------------
' TrackmaxForm.vba
' Christopher Ryan Bauer
' Version 0.1.0
'
' This document is used for an employee to request a new rebate program be
' created or modified. When distributed to employees protection needs to be
' turned on. This prevent them from doing anything but filling in the forms.
' We are using various procedures to ensure that the user input is valid, as
' well as turning on and off fields to make sure only the needed information is
' provided.
'
'   Private Properties;
'       blnInitialized:
'       blnPrintHiddenText:
'       blnShowHiddenText:
'       blnTableGridlines:
'       blnProtectDoc:
'       strDocPassword:
'       blnUpdateTitle:
'       strSaveName:
'       strPlaceholder(35):
'       strPrintSpace(35):
'       appEvents:
'       enuAmountFormat:
'       enuTagID:
'
'   Private Subs;
'       Document_Open()
'           - Occurs once when you open the document.
'
'       Document_Close()
'           - Occurs once right before the document is closed.
'
'       Document_ContentControlOnEnter(ByVal ContentControl As ContentControl)
'           - Triggers when the user enters a control.
'           ContentControl: The control that we are entering.
'
'       Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
'           - Triggers when the user leaves a control.
'           ContentControl: The control that was just left.
'           Cancel: If the event should be canceled and hold the user in the control.
'
'       appEvents_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
'           - Event occurring before the user saves.
'           Doc:        The document that is being saved.
'           SaveAsUI:   If they are doing a save or save as.
'           Cancel:     If set to true, cancels saving the document.
'
'       appEvents_DocumentBeforePrint(ByVal Doc As Document, Cancel As Boolean)
'           - Event occurring before the user prints.
'           Doc:    The document that is being printed.
'           Cancel: If set to true, cancels printing the document.
'
'       CaptureAppEvents()
'           - Ensure our appEvents object stays set.
'
'       InitializeDocument(Optional blnProtect As Boolean = False, Optional strPassword As String = "")
'           - Called when document is opened to save user settings.
'           blnProtect:     Used to decide if the document should be protected.
'           strPassword:    The password that will be used for protection.
'
'       InitializePlaceholders()
'           - Fills an array with placeholder values.
'
'       InitializePrintSpace()
'           - Fills an array with print spaces values.
'
'       ReleaseDocument()
'           - Called when document is closed to reset user settings.
'
'       ProtectDocument()
'           - Checks for protection, and if needed, protects the document.
'
'       UnProtectDocument()
'           - Checks for protection, and if needed removes it.
'
'       SetPlaceholders()
'           - Sets all of our user input fields placeholder text.
'
'       SetPrintLines(Optional LinesOn As Boolean = True)
'           - Creates lines for user to write on, or turns them off.
'           LinesOn:    True if we want to turn on lines, false if we want them off.
'
'       ResetList(strTag As String, Optional strDefaultText As String = "", Optional blnForceReset As Boolean = False)
'           - Resets a list control to its placeholder text based on tag.
'           strTag:         Will select all controls with this tag.
'           strDefaultText: The text we will change the list to.
'           blnForceReset:  Will force the subroutine to reset the text.
'
'       UpdateTitle(Optional blnForceUpdate As Boolean = False)
'           - Updates the documents title and suggested save name.
'           blnForceUpdate: Will force the subroutine to recalculate the title.
'
'       AddRSCLine(occSection As ContentControl)
'           - Adds a new line to the repeating section of the object passed.
'           occSection: The section containing a RSC we want a new line on.
'
'       UnlockRSC(tblSection As Table)
'           - Unlocks all but the first row of controls in an RSC.
'           tblSection: The main table object of the RSC we want to unlock.
'
'       SelectFirstEmptyRSC(tblSection As Table)
'           - Selects the first empty control in the passed RSC.
'           tblSection: The main table object of the RSC select a control in.
'
'       ClearRSC(rscSection As RepeatingSectionItemColl)
'           - Checks the passed RSC for rows with default text and removes them.
'           rscSection: The repeating section we want to clean up.
'
'       AccessIndex(intRow As Integer, intCell As Integer, intRows As Integer, intCells As Integer) As Integer
'           - Returns the index we need to use to access a control item.
'           intRow:     The index of the row we are on.
'           intCell:    The index of the cell we are on for this row.
'           intRows:    How many row are in teh table.
'           intCells:   How many cells are in each row.
'           Returns:    Integer representing what index to use to access the object.
'
'       GetRow(strID As String, strSection As String) As Integer
'           - Searches for the ID in the passed section and returns the row its on.
'           strID:      Will select all controls with this tag.
'           strSection: The text we will change the list to.
'           Returns:    Integer representing row of table ID was found on, 0 if not found.
'
'       GetEnumAmountFormatValue(strSymbol As String) As enuAmountFormat
'           - Returns the enum value of a passed symbol.
'           strSymbol:  The symbol of the enum value we want.
'           Returns:    The enum value for the symbol.
'
'       GetEnumTagIDValue(strTag As String) As enuTagID
'           - Returns the enum TagID value of the passed Tag.
'           strTag:     The tag as a string of the enum value we want.
'           Returns:    The enum value for the string.
'
'       GetEnumTagIDString(enuTagID As enuTagID) As String
'           - Returns the string of the enum TagID passed.
'           enuTagID:   The tag ID of the enum we want the string for.
'           Returns:    The string of the passed enum.
'
'       IsAlphaNumeric(strStatment As String, Optional blnAllowSpaces As Boolean = False) As Boolean
'           - Check to see if a statement is alphanumeric or not.
'           strStatment:    The statement we are checking.
'           blnAllowSpaces: If spaces should count as alphanumeric or not.
'           Returns:        True if all charters were alphanumeric, False otherwise.
'
'       IsPlaceholder(occTarget As ContentControl) As Boolean
'           - Checks to see if a control is set to its placeholder text.
'           occTarget:  The user input control that needs checked.
'           Returns:    True if text is same as placeholder, otherwise False.
'
'       IsPrintSpace(occTarget As ContentControl) As Boolean
'           - Checks to see if a control is set to its print spaces.
'           occTarget:  The user input control that needs checked.
'           Returns:    True if text is same as print space, otherwise False.
'
'       ValidateAmount(occTarget As ContentControl, Optional enuFormat As enuAmountFormat = RationalNumber) As Boolean
'           - Checks a $ or % ranges value to ensure good data and format.
'           occTarget:  The user input control that needs validated.
'           strFormat:  Which way to format range's text.
'           Returns:    True if data passed validation, False if an error was thrown.
'
'       AutofillCmbRequestedBy()
'           - Fills in the requester from the buyer.
'
'       AutofillLstBuyer()
'           - Fills in the buyer from the requester if it is a buyer.
'
'       AutofillTxtPayingVendor()
'           - Fills in the paying vendor from offering vendor.
'
'       AutofillTxtItemAmount(occItem As ContentControl)
'           - Fills in a default payment and amount for an item.
'           occItem: The content control whose user input we want checked.
'
'       CheckTxtItem(occItem As ContentControl) As Boolean
'           - Checks user enter data to make sure it is valid.
'           occItem: The content control whose user input we want checked.
'           Returns: Boolean showing if user input passed validation or not.
'
'       CheckTxtItemAmount(occItemAmount As ContentControl) As Boolean
'           - Checks user enter data to make sure it is valid.
'           occItemAmount: The content control whose user input we want checked.
'           Returns: Boolean showing if user input passed validation or not.
'
'       CheckTxtSupplier(occSupplier As ContentControl) As Boolean
'           - Checks user enter data to make sure it is valid.
'           occItem: The content control whose user input we want checked.
'           Returns: Boolean showing if user input passed validation or not.
'
'       CheckTxtCustomer(occCustomer As ContentControl) As Boolean
'           - Checks user enter data to make sure it is valid.
'           occItem: The content control whose user input we want checked.
'           Returns: Boolean showing if user input passed validation or not.
'
'       ToggelChkNewProgram()
'           - Turns on or off features based on the New Program field.
'
'       ToggelTxtNewProgramNumber()
'           - Turns on or off features based on being completed.
'
'       ToggelChkTrackingOnly()
'           - Turns on or off features based on chkTrackingOnly.
'
'       ToggelChkSales()
'           - Turns on or off features based on sales and other fields.
'
'       ToggelAllowChkRestrictNOI()
'           - Checks to see if NOI should be on or off.
'
'       ToggelChkOtherProgramType()
'           - This toggles other program type fields on or off.
'
'       ToggelLstCategory()
'           - This toggles the if other program category on and off.
'
'       ToggelChkEmailInvoice()
'           - This turns on or off our email fields.
'
'       ToggelChkPayBuyingGroup()
'           - This turns on or off our buying group fields.
'
'       ToggelLstRebateType()
'           - This turns on or off our buying group fields.
'
'       ToggelChkRestrictNOI()
'           - This turns on or off the Restrict NOI fields.
'
'       ToggelChkOnlyNewCustomers()
'           - This turns on or off new customers only fields.

Private blnInitialized As Boolean
Private blnPrintHiddenText As Boolean
Private blnShowHiddenText As Boolean
Private blnTableGridlines As Boolean
Private blnProtectDoc As Boolean
Private strDocPassword As String
Private blnUpdateTitle As Boolean
Private strSaveName As String
Private strPlaceholder(35) As String
Private strPrintSpace(35) As String
Private WithEvents appEvents As Word.Application

Private Enum enuAmountFormat
    WholeNumber
    RationalNumber
    Cash
    Percentage
End Enum

Private Enum enuTagID
    'Fields using placeholder text.
    txtProgramNumber
    dtpRequested
    cmbRequestedBy
    dtpCompleted
    cmbCompletedBy
    txtProgramTitle
    txtNewProgramNumber
    txtOfferingVendor
    txtOtherProgramType
    lstBuyer
    lstPayment
    txtDescription
    dtpStartDate
    dtpEndDate
    lstCategory
    txtOtherCategory
    txtPayingVendor
    txtToEmail1
    txtToEmail2
    txtToEmail3
    lstEmailFormat
    txtBuyingGroup
    lstGLAccount
    txtMfcProgramID
    lstRebateType
    txtAmount
    lstPaidBy
    txtNOIAmount
    lstNOIUnit
    txtNotes
    txtItem
    txtItemAmount
    lstItemPaidBy
    txtSupplier
    txtCustomer
    dtpNewSince
    
    'Fields not using placeholder text.
    lblProgramRequest
    chkNewProgram
    lblChkNewProgram
    chkModifyExisting
    lblChkModifyExisting
    lblTxtProgramNumber
    lblDtpRequested
    lblCmbRequestedBy
    lblDtpCompleted
    lblCmbCompletedBy
    lblTxtNewProgramNumber
    lblTxtOfferingVendor
    chkTrackingOnly
    lblChkTrackingOnly
    chkSales
    lblChkSales
    chkPurchasing
    lblChkPurchasing
    chkOtherProgramType
    lblChkOtherProgramType
    lblTxtOtherProgramType
    lblLstBuyer
    lblLstPayment
    lblTxtDescription
    lblDtpStartDate
    lblDtpEndDate
    chkInvoiced
    lblChkInvoiced
    chkReceived
    lblChkReceived
    chkOrdered
    lblChkOrdered
    lblLstCategory
    lblTxtOtherCategory
    lblTxtPayingVendor
    chkEmailInvoice
    lblChkEmailInvoice
    lblTxtToEmail1
    lblTxtToEmail2
    lblTxtToEmail3
    lblLstEmailFormat
    chkPayBuyingGroup
    lblChkPayBuyingGroup
    lblTxtBuyingGroup
    lblLstGLAccount
    lblTxtMfcProgramID
    lblLstRebateType
    lblTxtAmount
    lblLstPaidBy
    chkRestrictNOI
    lblChkRestrictNOI
    lblTxtNOIAmount
    lblLstNOIUnit
    lblTxtNotes
    lblRscEligibleItems
    txtLblBtnAddItems
    txtLblBtnClearItems
    rscEligibleItems
    lblRscEligibleSuppliers
    txtLblBtnAddSuppliers
    txtLblBtnClearSuppliers
    rscEligibleSuppliers
    lblRscEligibleCustomers
    txtLblBtnAddCustomers
    txtLblBtnClearCustomers
    rscEligibleCustomers
    chkOnlyNewCustomers
    lblChkOnlyNewCustomers
    lblDtpNewSince
End Enum
  
  
' ------------------------------------------------------------------------------
' Document_Open - Occurs once when you open the document.

Private Sub Document_Open()
    Call InitializeDocument(True, "MysticAlly@1025")
End Sub


' ------------------------------------------------------------------------------
' Document_Close - Occurs once right before the document is closed.

Private Sub Document_Close()
    ReleaseDocument
End Sub


' ------------------------------------------------------------------------------
' Document_ContentControlOnEnter - Triggers when the user enters a control.
'
' This event occurs whenever a users focus enters a content control. The control
' that was just entered is passed to use. We use a select on the controls .tag
' to decided what actions to perform. We only want to do simple variable
' assignments and sub/function calls here to make it easier to tell what each
' case is doing. Mos often a check box is being entered and we are turning off
' other in its group and toggling on/off fields. We are also using some text
' controls to act like buttons by calling function on their enter and then
' setting focus away from them so they can be "clicked" (re entered) repeatedly.
'
'   Parameters;
'       ContentControl: The control that we are entering.

Private Sub Document_ContentControlOnEnter(ByVal ContentControl As ContentControl)
    Call InitializeDocument(True, "MysticAlly@1025")
    CaptureAppEvents
    
    Select Case ContentControl.Tag
        Case "chkNewProgram"
            ActiveDocument.SelectContentControlsByTag("chkModifyExisting")(1).Checked = False
            ToggelChkNewProgram
            
        Case "chkModifyExisting"
            ActiveDocument.SelectContentControlsByTag("chkNewProgram")(1).Checked = False
            ToggelChkNewProgram
            
        Case "chkTrackingOnly"
            ToggelChkTrackingOnly
            
        Case "chkSales"
            ActiveDocument.SelectContentControlsByTag("chkPurchasing")(1).Checked = False
            ActiveDocument.SelectContentControlsByTag("chkOtherProgramType")(1).Checked = False
            ToggelChkSales
        
        Case "chkPurchasing"
            ActiveDocument.SelectContentControlsByTag("chkSales")(1).Checked = False
            ActiveDocument.SelectContentControlsByTag("chkOtherProgramType")(1).Checked = False
            ToggelChkSales
        
        Case "chkOtherProgramType"
            ActiveDocument.SelectContentControlsByTag("chkPurchasing")(1).Checked = False
            ActiveDocument.SelectContentControlsByTag("chkSales")(1).Checked = False
            ToggelChkSales
            
        Case "chkInvoiced"
            ActiveDocument.SelectContentControlsByTag("chkReceived")(1).Checked = False
            ActiveDocument.SelectContentControlsByTag("chkOrdered")(1).Checked = False
            ActiveDocument.SelectContentControlsByTag("lstCategory")(1).Range.Select
            
        Case "chkReceived"
            ActiveDocument.SelectContentControlsByTag("chkInvoiced")(1).Checked = False
            ActiveDocument.SelectContentControlsByTag("chkOrdered")(1).Checked = False
            ActiveDocument.SelectContentControlsByTag("lstCategory")(1).Range.Select
            
        Case "chkOrdered"
            ActiveDocument.SelectContentControlsByTag("chkReceived")(1).Checked = False
            ActiveDocument.SelectContentControlsByTag("chkInvoiced")(1).Checked = False
            ActiveDocument.SelectContentControlsByTag("lstCategory")(1).Range.Select
            
        Case "chkEmailInvoice"
            ToggelChkEmailInvoice
            
        Case "chkPayBuyingGroup"
            ToggelChkPayBuyingGroup
            
        Case "chkRestrictNOI"
            ToggelChkRestrictNOI
            
        Case "txtLblBtnAddItems"
            Call AddRSCLine(ActiveDocument.SelectContentControlsByTag("rscEligibleItems")(1))
            ToggelLstRebateType
            Call SelectFirstEmptyRSC(ActiveDocument.SelectContentControlsByTag("rscEligibleItems")(1).RepeatingSectionItems(1).Range.Tables(1))
             
        Case "txtLblBtnClearItems"
            Call ClearRSC(ActiveDocument.SelectContentControlsByTag("rscEligibleItems")(1).RepeatingSectionItems)
            Call SelectFirstEmptyRSC(ActiveDocument.SelectContentControlsByTag("rscEligibleItems")(1).RepeatingSectionItems(1).Range.Tables(1))
            
        Case "txtLblBtnAddSuppliers"
            Call AddRSCLine(ActiveDocument.SelectContentControlsByTag("rscEligibleSuppliers")(1))
            Call SelectFirstEmptyRSC(ActiveDocument.SelectContentControlsByTag("rscEligibleSuppliers")(1).RepeatingSectionItems(1).Range.Tables(1))
             
        Case "txtLblBtnClearSuppliers"
            Call ClearRSC(ActiveDocument.SelectContentControlsByTag("rscEligibleSuppliers")(1).RepeatingSectionItems)
            Call SelectFirstEmptyRSC(ActiveDocument.SelectContentControlsByTag("rscEligibleSuppliers")(1).RepeatingSectionItems(1).Range.Tables(1))
            
        Case "txtLblBtnAddCustomers"
            If Not ActiveDocument.SelectContentControlsByTag("chkOtherProgramType")(1).Checked _
            And Not ActiveDocument.SelectContentControlsByTag("chkPurchasing")(1).Checked Then
                Call AddRSCLine(ActiveDocument.SelectContentControlsByTag("rscEligibleCustomers")(1))
                Call SelectFirstEmptyRSC(ActiveDocument.SelectContentControlsByTag("rscEligibleCustomers")(1).RepeatingSectionItems(1).Range.Tables(1))
            End If
             
        Case "txtLblBtnClearCustomers"
            If Not ActiveDocument.SelectContentControlsByTag("chkOtherProgramType")(1).Checked _
            And Not ActiveDocument.SelectContentControlsByTag("chkPurchasing")(1).Checked Then
                Call ClearRSC(ActiveDocument.SelectContentControlsByTag("rscEligibleCustomers")(1).RepeatingSectionItems)
                Call SelectFirstEmptyRSC(ActiveDocument.SelectContentControlsByTag("rscEligibleCustomers")(1).RepeatingSectionItems(1).Range.Tables(1))
            End If
            
        Case "chkOnlyNewCustomers"
            ToggelChkOnlyNewCustomers
    End Select
End Sub


' ------------------------------------------------------------------------------
' Document_ContentControlOnExit - Triggers when the user leaves a control.
'
' This event occurs whenever a users focus leaves a content control. The control
' that was just left is passed to us. We use its .tag property tin a select to
' decide what functions we need to be performing. We only want to do simple
' variable assignment and function/sub calls here to make it easy to see what we
' are doing in each case. The most common events are that a check box was left
' and we need to call a toggle function to check the state and turn appropriate
' fields on or off. Or that a data field was left and we need to check and
' validate user input. In the case that we have bad input returned from a check
' we are using exit sub instead of setting cancel = to true so we don't update
' our title.
'
'   Parameters;
'       ContentControl: The control that was just left.
'       Cancel: If the event should be canceled and hold the user in the control.

Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
    Call InitializeDocument(True, "MysticAlly@1025")
    CaptureAppEvents
    
    Select Case ContentControl.Tag
        Case "cmbRequestedBy"
            AutofillLstBuyer
            
        Case "cmbCompletedBy"
            ToggelTxtNewProgramNumber
            
        Case "dtpCompleted"
            ToggelTxtNewProgramNumber
            
        Case "txtProgramNumber"
            If Not ValidateAmount(ContentControl, WholeNumber) Then: Exit Sub
            
        Case "txtNewProgramNumber"
            If Not ValidateAmount(ContentControl, WholeNumber) Then: Exit Sub
            
        Case "txtOfferingVendor"
            AutofillTxtPayingVendor
            If Not ValidateAmount(ContentControl, WholeNumber) Then Exit Sub
    
        Case "lstBuyer"
            blnUpdateTitle = True
            Call ResetList(ContentControl.Tag)
            AutofillCmbRequestedBy
            
        Case "lstPayment"
            blnUpdateTitle = True
            Call ResetList(ContentControl.Tag)
            
        Case "txtDescription"
            blnUpdateTitle = True
            
        Case "lstCategory"
            blnUpdateTitle = True
            Call ResetList(ContentControl.Tag)
            ToggelLstCategory
            
        Case "txtOtherCategory"
            blnUpdateTitle = True
            
        Case "txtPayingVendor"
            If Not ValidateAmount(ContentControl, WholeNumber) Then Exit Sub
            
        Case "lstEmailFormat"
            Call ResetList(ContentControl.Tag)
        
        Case "lstGLAccount"
            Call ResetList(ContentControl.Tag)
        
        Case "lstRebateType"
            Call ResetList(ContentControl.Tag)
            ToggelLstRebateType
            
        Case "txtAmount"
            If Not ValidateAmount(ContentControl, GetEnumAmountFormatValue(Left(ActiveDocument.SelectContentControlsByTag("lstPaidBy")(1).Range.Text, 1))) Then: Exit Sub
            
        Case "lstPaidBy"
            Call ResetList(ContentControl.Tag)
            If Not ValidateAmount(ActiveDocument.SelectContentControlsByTag("txtAmount")(1), GetEnumAmountFormatValue(Left(ContentControl.Range.Text, 1))) Then Exit Sub
         
        Case "txtNOIAmount"
            If Not ValidateAmount(ContentControl, Cash) Then: Exit Sub
            
        Case "lstNOIUnit"
            Call ResetList(ContentControl.Tag)
            
        Case "txtItem"
            If Not CheckTxtItem(ContentControl) Then: Exit Sub
            Call AutofillTxtItemAmount(ContentControl)
            
        Case "txtItemAmount"
            If Not CheckTxtItemAmount(ContentControl) Then: Exit Sub
        
        Case "lstItemPaidBy"
            Call ResetList(ContentControl.Tag)
            If Not CheckTxtItemAmount(ContentControl.ParentContentControl.RepeatingSectionItems.Item(1).Range.Tables(1).Rows(GetRow( _
                ContentControl.ID, ContentControl.ParentContentControl.Tag)).Cells(2).Range.ContentControls(1)) Then: Exit Sub
            
        Case "txtSupplier"
            If Not CheckTxtSupplier(ContentControl) Then: Exit Sub
            
        Case "txtCustomer"
            If Not CheckTxtCustomer(ContentControl) Then: Exit Sub
    End Select
    
    UpdateTitle
End Sub


' ------------------------------------------------------------------------------
' appEvents_DocumentBeforeSave - Event occurring before the user saves.
'
' This is a document event trigger by the user selecting to save the document.
' We are capturing it with our appEvents object in order to assist the user by
' suggesting a document name if they are doing a save as. We must call our own
' save procedure to do this, so we need to cancel the one they invoked.
'
' Parameters;
'       Doc:        The document that is being saved.
'       SaveAsUI:   If they are doing a save or save as.
'       Cancel:     If set to true, cancels saving the document.

Private Sub appEvents_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
    If SaveAsUI Then
      With Dialogs(wdDialogFileSaveAs)
        .Name = ActiveDocument.Path & "\" & strSaveName
        .Format = Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled
        .Show
      End With
      
      Cancel = True
    End If
End Sub

 
' ------------------------------------------------------------------------------
' appEvents_DocumentBeforePrint - Event occurring before the user prints.
'
' This is a document event trigger by the user selecting to print the document.
' We are capturing it with our appEvents object in order to replace any
' placeholder text with blank lines to write on. Because there is no after print
' event, we can not set our lines, allow a normal print to happen, and then set
' the placeholder back. So we will need to cancel the print event that the user
' called, and prompt our own dialog in between turning the lines on and off.
'
' Parameters;
'       Doc:    The document that is being printed.
'       Cancel: If set to true, cancels printing the document.

Private Sub appEvents_DocumentBeforePrint(ByVal Doc As Document, Cancel As Boolean)
    Cancel = True
    SetPrintLines
    Dialogs(wdDialogFilePrint).Show
    SetPrintLines (False)
End Sub


' ------------------------------------------------------------------------------
' CaptureAppEvents - Ensure our appEvents object stays set.
'
' Due to the appEvents object frequently becoming unset. Most likely due to
' quitting running code while debugging. We have this event to make sure that
' appEvents is set, and if it is not then set it. This probably is not needed
' once the document is released, but it should not do any harm to keep it in
' case appEvents becomes unset. Place it somewhere where it will be called
' periodically, like in the on enter/exit subs.

Private Sub CaptureAppEvents()
    If appEvents Is Nothing Then: Set appEvents = Word.Application
End Sub


' ------------------------------------------------------------------------------
' InitializeDocument - Called when document is opened to save user settings.
'
' This subroutine saves some of the users preferences before altering them so
' that our document will display properly. While ensuring we are able to set the
' users preferences back to normal before we close. It sets our appEvents so that
' we are able to capture word application events. It also ensures that our doc
' is properly protected. It should only be run once or we will overwrite the
' users preferences. To avoid this we will check the blnInitialized first.
'
'   Parameters;
'       blnProtect:     Used to decide if the document should be protected.
'       strPassword:    The password that will be used for protection.

Private Sub InitializeDocument(Optional blnProtect As Boolean = False, Optional strPassword As String = "")
    If Not blnInitialized Then
        blnInitialized = True
        blnProtectDoc = blnProtect
        strDocPassword = strPassword
        blnPrintHiddenText = Options.PrintHiddenText
        blnShowHiddenText = Application.ActiveWindow.View.ShowHiddenText
        blnTableGridlines = Application.ActiveWindow.View.TableGridlines
        strSaveName = "Incomplete Program Request " & Format(Date, "YYYYMMDD")
        blnUpdateTitle = False
        
        InitializePlaceholders
        SetPlaceholders
    
        Options.PrintHiddenText = False
        Application.ActiveWindow.View.ShowHiddenText = True
        Application.ActiveWindow.View.TableGridlines = False
        
        CaptureAppEvents
        
        ProtectDocument
    End If
End Sub


' ------------------------------------------------------------------------------
' InitializePlaceholders - Fills an array with placeholder values.
'
' This subroutine fills an array with placeholder text for each user modifiable
' field in the form. THis is used when checking to see if a field is in a
' default state, or when reseting to a default state.

Private Sub InitializePlaceholders()
    strPlaceholder(dtpRequested) = "Choose or type a date"
    strPlaceholder(cmbRequestedBy) = "Choose or type a name"
    strPlaceholder(dtpCompleted) = "Choose or type a date"
    strPlaceholder(cmbCompletedBy) = "Choose or type a name"
    
    strPlaceholder(txtProgramNumber) = "Type number"
    strPlaceholder(txtProgramTitle) = " "
    strPlaceholder(txtNewProgramNumber) = "Enter program #"
    
    strPlaceholder(txtOfferingVendor) = "Enter vendor #"
    strPlaceholder(txtOtherProgramType) = "Enter program type"
    strPlaceholder(lstBuyer) = "Select a buyer"
    strPlaceholder(lstPayment) = "Select a payment type"
    strPlaceholder(txtDescription) = "Enter a program description"
    strPlaceholder(dtpStartDate) = "Choose or type a date"
    strPlaceholder(dtpEndDate) = "Choose or type a date"
    strPlaceholder(lstCategory) = "Select a category"
    strPlaceholder(txtOtherCategory) = "Enter category"
    
    strPlaceholder(txtPayingVendor) = "Enter vendor #"
    strPlaceholder(txtToEmail1) = "Enter email"
    strPlaceholder(txtToEmail2) = "Enter email (optional)"
    strPlaceholder(txtToEmail3) = "Enter email (optional)"
    strPlaceholder(lstEmailFormat) = "Select format"
    strPlaceholder(txtBuyingGroup) = "Enter group"
    strPlaceholder(lstGLAccount) = "Select account"
    strPlaceholder(txtMfcProgramID) = "Enter ID"
    
    strPlaceholder(lstRebateType) = "Select type"
    strPlaceholder(txtAmount) = "Enter amount (optional)"
    strPlaceholder(lstPaidBy) = "Select payment"
    strPlaceholder(txtNOIAmount) = "Enter amount"
    strPlaceholder(lstNOIUnit) = "Select unit"
    
    strPlaceholder(txtNotes) = "Enter any additional notes (optional)"
    
    strPlaceholder(txtItem) = "Enter item #"
    strPlaceholder(txtItemAmount) = "Enter amount"
    strPlaceholder(lstItemPaidBy) = "Select payment"
    strPlaceholder(txtSupplier) = "Enter supplier #"
    strPlaceholder(txtCustomer) = "Enter customer #"
    strPlaceholder(dtpNewSince) = "Choose or type a date"
End Sub


' ------------------------------------------------------------------------------
' InitializePrintSpace - Fills an array with print spaces values.
'
' This subroutine fills an array with spaces for each user editable field.
' This is used before printing, we change any placeholders to a set number of
' spaces and turn on underline to make lines to write on.

Private Sub InitializePrintSpace()
    strPrintSpace(dtpRequested) = String(29, "_")
    strPrintSpace(cmbRequestedBy) = String(36, "_")
    strPrintSpace(dtpCompleted) = String(29, "_")
    strPrintSpace(cmbCompletedBy) = String(36, "_")
    
    strPrintSpace(txtProgramNumber) = String(14, "_")
    strPrintSpace(txtProgramTitle) = String(36, "_")
    strPrintSpace(txtNewProgramNumber) = String(26, "_")
    
    strPrintSpace(txtOfferingVendor) = String(23, "_")
    strPrintSpace(txtOtherProgramType) = String(33, "_")
    strPrintSpace(lstBuyer) = String(33, "_")
    strPrintSpace(lstPayment) = String(26, "_")
    strPrintSpace(txtDescription) = String(29, "_")
    strPrintSpace(dtpStartDate) = String(30, "_")
    strPrintSpace(dtpEndDate) = String(31, "_")
    strPrintSpace(lstCategory) = String(31, "_")
    strPrintSpace(txtOtherCategory) = String(33, "_")
    
    strPrintSpace(txtPayingVendor) = String(25, "_")
    strPrintSpace(txtToEmail1) = String(37, "_")
    strPrintSpace(txtToEmail2) = String(38, "_")
    strPrintSpace(txtToEmail3) = String(38, "_")
    strPrintSpace(lstEmailFormat) = String(33, "_")
    strPrintSpace(txtBuyingGroup) = String(28, "_")
    strPrintSpace(lstGLAccount) = String(29, "_")
    strPrintSpace(txtMfcProgramID) = String(25, "_")
    
    strPrintSpace(lstRebateType) = String(28, "_")
    strPrintSpace(txtAmount) = String(26, "_")
    strPrintSpace(lstPaidBy) = String(33, "_")
    strPrintSpace(txtNOIAmount) = String(33, "_")
    strPrintSpace(lstNOIUnit) = String(36, "_")
    
    strPrintSpace(txtNotes) = String(32, "_") & " " & String(38, "_") & " " & String(38, "_") & " " & String(38, "_")
    
    strPrintSpace(txtItem) = String(13, "_")
    strPrintSpace(txtItemAmount) = String(13, "_")
    strPrintSpace(lstItemPaidBy) = String(13, "_")
    strPrintSpace(txtSupplier) = String(18, "_")
    strPrintSpace(txtCustomer) = String(18, "_")
    strPrintSpace(dtpNewSince) = String(21, "_")
End Sub


' ------------------------------------------------------------------------------
' ReleaseDocument - Called when document is closed to reset user settings.
'
' This subroutine uses the variable we saved when initializing the document to
' reset the user settings back to what they were before using our document.

Private Sub ReleaseDocument()
    If blnInitialized Then
        blnInitialized = False
        Options.PrintHiddenText = blnPrintHiddenText
        ActiveWindow.View.ShowHiddenText = blnShowHiddenText
        ActiveWindow.View.TableGridlines = blnTableGridlines
    End If
End Sub


' ------------------------------------------------------------------------------
' ProtectDocument - Checks for protection, and if needed, protects the document.
'
' Uses the document variable blnProtectDoc to decide if the document should be
' protected or not. Check to see what state the document is in before attempting
' to protect or un-protect it as not doing so will cause an error. Uses the
' document variable strDocPassword as the password for protection.

Private Sub ProtectDocument()
    If blnProtectDoc Then
        If ActiveDocument.ProtectionType = wdNoProtection Then: Call ActiveDocument.Protect(wdAllowOnlyFormFields, True, strDocPassword)
    Else
        If ActiveDocument.ProtectionType <> wdNoProtection Then: Call ActiveDocument.Unprotect(strDocPassword)
    End If
End Sub


' ------------------------------------------------------------------------------
' UnProtectDocument - Checks for protection, and if needed removes it.
'
' Check to see if the document is protected, and if so uses the strDocPassword
' to un-protect the document. Should be called before making dynamic changes such
' as adding and removing repeating section items

Private Sub UnProtectDocument()
    If ActiveDocument.ProtectionType <> wdNoProtection Then: Call ActiveDocument.Unprotect(strDocPassword)
End Sub


' ------------------------------------------------------------------------------
' SetPlaceholders - Sets all of our user input fields placeholder text.
'
' This subroutine loops through all of the user input fields and uses the array
' of placeholder strings to set the placeholders for each content control.

Private Sub SetPlaceholders()
    Dim occTarget As ContentControl
    Dim enuTag As enuTagID
    
    For enuTag = txtProgramNumber To dtpNewSince
        Set occTarget = ActiveDocument.SelectContentControlsByTag(GetEnumTagIDString(enuTag))(1)
        Call occTarget.SetPlaceholderText(, , strPlaceholder(enuTag))
    Next enuTag
End Sub


' ------------------------------------------------------------------------------
' SetPrintLines - Creates lines for user to write on, or turns them off.
'
' This is a document event trigger by the user selecting to print the document.
' We are capturing it with our appEvents object in order to replace any
' placeholder text with blank lines to write on. Because there is no after print
' event, we can not set our lines, allow a normal print to happen, and then set
' the placeholder back. So we will need to cancel the print event that the user
' called, and prompt our own dialog in between turning the lines on and off.
'
' Parameters;
'       LinesOn:    True if we want to turn on lines, false if we want them off.

Private Sub SetPrintLines(Optional LinesOn As Boolean = True)
    Dim occTarget As ContentControl
    Dim enuTag As enuTagID
    Dim enuType As WdContentControlType
    Dim blnLock As Boolean

    For Each occTarget In ActiveDocument.ContentControls
        enuTag = GetEnumTagIDValue(occTarget.Tag)
        If enuTag <= dtpNewSince Then
            If LinesOn And IsPlaceholder(occTarget) Or Not LinesOn And IsPrintSpace(occTarget) Then
                blnLock = occTarget.LockContents
                occTarget.LockContents = False
                enuType = occTarget.Type
                If enuType <> wdContentControlText Then: occTarget.Type = wdContentControlText
                
                If LinesOn Then
                    occTarget.Range.Text = strPrintSpace(enuTag)
                Else
                    occTarget.Range.Text = ""
                End If
                
                If enuType <> wdContentControlText Then: occTarget.Type = enuType
                occTarget.LockContents = blnLock
            End If
        End If
    Next occTarget
End Sub


' ------------------------------------------------------------------------------
' ResetList - Resets a list control to its placeholder text based on tag.
'
' This function will select all list controls with the passed tag and if they
' are blank(One Space) reset them to null so they pick up their placeholder. A
' string can be passed to be used instead of the placeholder text. The list can
' be forced to be reset even if it is not blank as well.
'
'   Parameters;
'       strTag:         Will select all controls with this tag.
'       strDefaultText: The text we will change the list to.
'       blnForceReset:  Will force the subroutine to reset the text.

Private Sub ResetList(strTag As String, Optional strDefaultText As String = "", Optional blnForceReset As Boolean = False)
    Dim ocsLists As ContentControls
    Dim occList As ContentControl
    
    Set ocsLists = ActiveDocument.SelectContentControlsByTag(strTag)
    
    For Each occList In ocsLists
        With occList
        If .Type = wdContentControlDropdownList _
        And .Range.Text = " " _
        Or blnForceReset Then
                    .LockContents = False
                    .Type = wdContentControlText
                    .Range.Text = strDefaultText
                    .Type = wdContentControlDropdownList
                End If
        End With
    Next occList
End Sub


' ------------------------------------------------------------------------------
' UpdateTitle - Updates the documents title and suggested save name.
'
' Check to see if any changes have been made that would effect the title. If so
' processes the forms vales into a new title. Updates the title in the document,
' its properties, the caption, and the suggested save name.
'
'   Parameters;
'       blnForceUpdate: Will force the subroutine to recalculate the title.

Private Sub UpdateTitle(Optional blnForceUpdate As Boolean = False)
    If blnUpdateTitle Or blnForceUpdate Then
        Dim strTitle As String
        Dim strCaption As String
        Dim strCategory As String
        Dim strOtherCategory As String
        Dim strDescription As String
        Dim strBuyer As String
        Dim strPayment As String
        Dim strDate As String
        
        strCategory = ActiveDocument.SelectContentControlsByTag("lstCategory")(1).Range.Text
        strOtherCategory = ActiveDocument.SelectContentControlsByTag("txtOtherCategory")(1).Range.Text
        strDescription = ActiveDocument.SelectContentControlsByTag("txtDescription")(1).Range.Text
        strBuyer = ActiveDocument.SelectContentControlsByTag("lstBuyer")(1).Range.Text
        strPayment = ActiveDocument.SelectContentControlsByTag("lstPayment")(1).Range.Text
        strDate = ActiveDocument.SelectContentControlsByTag("dtpStartDate")(1).Range.Text
        
        Select Case strCategory
            Case " ", "Select a category"
                strCategory = ""
                
            Case "Other"
                Select Case strOtherCategory
                    Case " ", "Enter category"
                        strCategory = ""
                        
                    Case Else
                        strCategory = strOtherCategory
                End Select
        End Select

        If strDescription = " " Or strDescription = "Enter a program description" Then
            strDescription = ""
        End If
            
        Select Case strBuyer
            Case "Amy Carson":      strBuyer = "21"
            Case "Greg Mason":      strBuyer = "8"
            Case "Jim Kelly":       strBuyer = "15"
            Case "Karen Stevenson": strBuyer = "3"
            Case "Kevin Mason":     strBuyer = "6"
            Case "Sandy Hilton":    strBuyer = "14"
            Case "Tom Wuerth":      strBuyer = "1"
            Case "Tracy Kissel":    strBuyer = "4"
            Case Else:              strBuyer = ""
        End Select
        
        Select Case strPayment
            Case "Check":           strPayment = "CHK"
            Case "Deduction":       strPayment = "DED"
            Case "EDA":             strPayment = "EDA"
            Case "Tracking Only":   strPayment = "TRK"
            Case Else:              strPayment = ""
        End Select
        
        If IsDate(strDate) Then
            strDate = Format(strDate, "YYYY-MM-DD")
        Else
            strDate = Format(Date, "YYYY-MM-DD")
        End If
        
        If strCategory = "" Or strDescription = "" Or strBuyer = "" Or strPayment = "" Then
            strTitle = ""
            strCaption = "Incomplete Program Request"
        Else
            strTitle = strCategory & ": " & strDescription & " (" & strBuyer & "-" & strPayment & ")"
            strCaption = strTitle
        End If
            
        With ActiveDocument.SelectContentControlsByTag("txtProgramTitle")(1)
            .LockContents = False
            .Range.Text = strTitle
            .LockContents = True
        End With
            
        ActiveDocument.ActiveWindow.Caption = strCaption
        strSaveName = strCaption & " " & strDate
        
        blnUpdateTitle = False
    End If 'blnUpdateTitle or blnForceUpdate
End Sub


' ------------------------------------------------------------------------------
' AddRSCLine - Adds a new line to the repeating section of the object passed.
'
' This subroutine unlocks the document and section passed, adds a new line after
' the last item in the list, and unlocks the newly added items. We then re-lock
' the section and document.
'
'   Parameters;
'       occSection: The section containing a RSC we want a new line on.

Private Sub AddRSCLine(occSection As ContentControl)
    UnProtectDocument
    occSection.LockContentControl = False
    occSection.LockContents = False
    occSection.AllowInsertDeleteSection = True
    
    occSection.RepeatingSectionItems.Item(occSection.Range.Rows.Count).InsertItemAfter
    Call UnlockRSC(occSection.RepeatingSectionItems.Item(1).Range.Tables(1))
    
    occSection.LockContentControl = True
    occSection.LockContents = True
    occSection.AllowInsertDeleteSection = False
    ProtectDocument
End Sub


' ------------------------------------------------------------------------------
' UnlockRSC - Unlocks all but the first row of controls in an RSC.
'
' This subroutine table a table of a repeating section, loops over each row
' except the first, and loops through each cell. It selects the content control
' in that cell and unlocks it so that it may be deleted when needed. We are not
' unlocking the first row because deleting all of the rows of a repeating
' section would completely remove the repeating section.
'
'   Parameters;
'       tblSection: The main table object of the RSC we want to unlock.

Private Sub UnlockRSC(tblSection As Table)
    Dim intRow As Integer
    Dim intCell As Integer
    
    For intRow = 2 To tblSection.Rows.Count
        For intCell = 1 To tblSection.Rows(intRow).Cells.Count
            tblSection.Rows(intRow).Cells(intCell).Range.ContentControls( _
            AccessIndex(intRow, intCell, tblSection.Rows.Count, tblSection.Rows(intRow).Cells.Count)) _
            .LockContentControl = False
        Next intCell
    Next intRow
End Sub


' ------------------------------------------------------------------------------
' SelectFirstEmptyRSC - Selects the first empty control in the passed RSC.
'
' This subroutine takes a table of a repeating section and search from the start
' for a content control that is empty. Once it is found it will select that
' control and exit the sub. If none is found it selects the last item.
'
'   Parameters;
'       tblSection: The main table object of the RSC select a control in.

Private Sub SelectFirstEmptyRSC(tblSection As Table)
    Dim intRow As Integer
    Dim intCell As Integer
    Dim intRows As Integer
    Dim intCells As Integer
    Dim intAccess As Integer
    
    intRows = tblSection.Rows.Count
    For intRow = 1 To intRows
        intCells = tblSection.Rows(intRow).Cells.Count
        For intCell = 1 To intCells
            intAccess = AccessIndex(intRow, intCell, intRows, intCells)
        
            If IsPlaceholder(tblSection.Rows(intRow).Cells(intCell).Range.ContentControls(intAccess)) Then
                tblSection.Rows(intRow).Cells(intCell).Range.ContentControls(intAccess).Range.Select
                Exit Sub
            End If
            
            If intRow = intRows And intCell = intCells Then: tblSection.Rows(intRow).Cells(intCell).Range.ContentControls(intAccess).Range.Select
        Next intCell
    Next intRow
End Sub


' ------------------------------------------------------------------------------
' ClearRSC - Checks the passed RSC for rows with default text and removes them.
'
' This subroutine may need to modify the document, so it will un and re protect
' the document and the RSC it is working on. It checks to see if all the inputs
' for a line in an RSC have been left with their placeholder text. If so it will
' remove those lines. But will always leave at lest one line.
'
' Due to buggy behavior when trying to remove the last row in an RSC we have
' to take some odd steps. We will be starting at the end of our RSC and working
' backwards until we have 2 rows left unchecked. Then we will see how may rows
' are still left, as its possible this subroutine was called with only 1 row
' existing. If we have more than 1 row we are going to check each one to see if
' it has user data on it or not, and store that in an array.
'
' We will use that data to decide what to do with the rows. We need to follow
' the rule that the last row can not be deleted. So if the last row dose not
' have user data on it, and there is data on 1 or 2, we will copy one of them to
' the last row, and delete the one we copied from.
'
' This will cause an issue when copying default text, as the document now thinks
' this was user entered. We will then have to check our fields to see if they
' have placeholder or disabled text. If they do we will need to remove the
' placeholder and see it the proper way. If the field is disabled, we need to
' make sure to set the proper text formating.
'
' Finally we will need to make sure to lock row 1 so it cant be deleted.
'
'   Parameters;
'       rscSection: The repeating section we want to clean up.

Private Sub ClearRSC(rscSection As RepeatingSectionItemColl)
    Dim occItem As ContentControl
    Dim intRow As Integer
    Dim intCell As Integer
    Dim intRows As Integer
    Dim intCells As Integer
    Dim intAccess As Integer
    Dim blnDelete As Boolean
    Dim blnDataOnRow(2) As Boolean
    Dim enuType As WdContentControlType

    UnProtectDocument
    rscSection.Parent.LockContentControl = False
    rscSection.Parent.LockContents = False
    rscSection.Parent.AllowInsertDeleteSection = True

    intRows = rscSection.Count
    For intRow = intRows To 3 Step -1
        blnDelete = True
        
        intCells = rscSection(intRow).Range.Cells.Count
        For intCell = 1 To intCells
            intAccess = AccessIndex(intRow, intCell, rscSection.Count, intCells)
            Set occItem = rscSection(intRow).Range.Cells(intCell).Range.ContentControls(intAccess)
            
            If Not IsPlaceholder(occItem) Or occItem.Range.Text = "Disabled" Then: blnDelete = False
        Next intCell
        
        If blnDelete Then: rscSection(intRow).Delete
    Next intRow
      
      
    intRows = rscSection.Count
    If intRows > 1 Then
    
        If intRows > 3 Then: intRows = 3
        For intRow = 1 To intRows
        
            intCells = rscSection(intRow).Range.Cells.Count
            For intCell = 1 To intCells
                intAccess = AccessIndex(intRow, intCell, rscSection.Count, intCells)
                Set occItem = rscSection(intRow).Range.Cells(intCell).Range.ContentControls(intAccess)
                
                If Not IsPlaceholder(occItem) Or occItem.Range.Text = "Disabled" Then: blnDataOnRow(intRow - 1) = True
            Next intCell
            
        Next intRow
    
    
        intRows = rscSection.Count
        intCells = rscSection(1).Range.Cells.Count
        For intCell = 1 To intCells
            intAccess = AccessIndex(1, intCell, intRows, intCells)
            Set occItem = rscSection(1).Range.Cells(intCell).Range.ContentControls(intAccess)
            occItem.LockContentControl = False
            occItem.LockContents = False
        Next intCell
        
        
        Select Case True
            Case Not blnDataOnRow(0) And Not blnDataOnRow(1) And Not blnDataOnRow(2)
                rscSection(1).Delete
            
            Case Not blnDataOnRow(0) And Not blnDataOnRow(1) And blnDataOnRow(2)
                rscSection(2).Delete
                rscSection(1).Delete
            
            Case Not blnDataOnRow(0) And blnDataOnRow(1) And Not blnDataOnRow(2)
                rscSection(1).Delete
                
            Case Not blnDataOnRow(0) And blnDataOnRow(1) And blnDataOnRow(2)
                rscSection(1).Delete
            
            Case blnDataOnRow(0) And Not blnDataOnRow(1) And Not blnDataOnRow(2)
                For intCell = 1 To intCells
                    intAccess = AccessIndex(1, intCell, intRows, intCells)
                    Set occItem = rscSection(1).Range.Cells(intCell).Range.ContentControls(intAccess)
                    
                    intAccess = AccessIndex(2, intCell, intRows, intCells)
                    With rscSection(2).Range.Cells(intCell).Range.ContentControls(intAccess)
                        .LockContents = False
                        enuType = .Type
                        .Type = wdContentControlText
                        .Range.Text = occItem.Range.Text
                        .Type = enuType
                    End With
                Next intCell
                
                rscSection(1).Delete
            
            Case blnDataOnRow(0) And Not blnDataOnRow(1) And blnDataOnRow(2)
                rscSection(2).Delete
        End Select
    
        
        intRows = rscSection.Count
        intCells = rscSection(1).Range.Cells.Count
        For intRow = 1 To intRows
            For intCell = 1 To intCells
                intAccess = AccessIndex(intRow, intCell, intRows, intCells)
                Set occItem = rscSection(intRow).Range.Cells(intCell).Range.ContentControls(intAccess)
                
                If IsPlaceholder(occItem) Then
                    occItem.LockContents = False
                    enuType = occItem.Type
                    occItem.Type = wdContentControlText
                    occItem.Range.Text = ""
                    occItem.Type = enuType
                ElseIf occItem.Range.Text = "Disabled" Then
                    occItem.LockContents = False
                    occItem.DefaultTextStyle = "OffList"
                    occItem.Appearance = wdContentControlHidden
                    occItem.LockContents = True
                End If
            Next intCell
        Next intRow
        
        
        intRows = rscSection.Count
        intCells = rscSection(1).Range.Cells.Count
        For intCell = 1 To intCells
            intAccess = AccessIndex(1, intCell, intRows, intCells)
            Set occItem = rscSection(1).Range.Cells(intCell).Range.ContentControls(intAccess)
            occItem.LockContentControl = True
        Next intCell
        
    End If 'intRows > 1
    
    rscSection.Parent.LockContentControl = True
    rscSection.Parent.LockContents = True
    rscSection.Parent.AllowInsertDeleteSection = False
    ProtectDocument
End Sub


' ------------------------------------------------------------------------------
' AccessIndex - Returns the index we need to use to access a control item.
'
' This function is used to deal with some odd behavior content controls have.
' When dealing with table or collections of content controls as a two dimensional
' table the item at the first index should be a single content control object.
' But if we are accessing the first(Row 1 Cell 1) or last item in the table,
' then the item at index 1 is a reference to the entire table. As such we need to
' access the item at index 2 to get our single object.
'
'   Parameters;
'       intRow:     The index of the row we are on.
'       intCell:    The index of the cell we are on for this row.
'       intRows:    How many row are in the table.
'       intCells:   How many cells are in each row.
'
'   Returns: Integer representing what index to use to access the object.

Private Function AccessIndex(intRow As Integer, intCell As Integer, intRows As Integer, intCells As Integer) As Integer
    If intRow = 1 And intCell = 1 Then
        AccessIndex = 2
        
    ElseIf intRow = intRows And intCell = intCells Then
        AccessIndex = 2
        
    Else
        AccessIndex = 1
      
    End If
End Function


' ------------------------------------------------------------------------------
' GetRow - Searches for the ID in the passed section and returns the row its on.
'
' This function checks each content controls ID in the passed section as a
' table. If it finds a match it returns the row of the table it found it on. If
' a match was not found it returns 0 to indicate it was missing.
'
'   Parameters;
'       strID:      Will select all controls with this tag.
'       strSection: The text we will change the list to.
'
'   Returns: Integer representing row of table ID was found on, 0 if not found.

Private Function GetRow(strID As String, strSection As String) As Integer
    Dim tblSection As Table
    Dim intRow As Integer
    Dim intCell As Integer
    Dim intRows As Integer
    Dim intCells As Integer

    Set tblSection = ActiveDocument.SelectContentControlsByTag(strSection)(1).RepeatingSectionItems.Item(1).Range.Tables(1)
    
    intRows = tblSection.Rows.Count
    For intRow = 1 To intRows
    
        intCells = tblSection.Rows(intRow).Cells.Count
        For intCell = 1 To intCells
        
            If strID = tblSection.Rows(intRow).Cells(intCell).Range.ContentControls.Item(AccessIndex(intRow, intCell, intRows, intCells)).ID Then
                GetRow = intRow
                Exit Function
            End If
        
        Next intCell
        
    Next intRow
    
    GetRow = 0
End Function


' ------------------------------------------------------------------------------
' GetEnumAmountFormatValue - Returns the enum value of a passed symbol.
'
' This function is needed because unlike .net, we can not use enum.parse to cast
' a string to an enum. As such we have a select case with each of our possible
' enum types represented by their symbol, which will return the corresponding
' enum type as a enuAmountFormat.
'
'   Parameters;
'       strSymbol: The symbol of the enum value we want.
'
'   Returns: The enum value for the symbol.

Private Function GetEnumAmountFormatValue(strSymbol As String) As enuAmountFormat
    Select Case strSymbol
        Case "$":       GetEnumAmountFormatValue = Cash
        Case "%":       GetEnumAmountFormatValue = Percentage
        Case "W", "w":  GetEnumAmountFormatValue = WholeNumber
        Case "R", "r":  GetEnumAmountFormatValue = RationalNumber
        Case Else:      GetEnumAmountFormatValue = RationalNumber
    End Select
End Function


' ------------------------------------------------------------------------------
' GetEnumTagIDValue - Returns the enum TagID value of the passed Tag.
'
' This function is needed because unlike .net, we can not use enum.parse to cast
' a string to an enum. As such we have a big case statement with all possible
' tags as a string retuning their enum.
'
'   Parameters;
'       strTag: The tag as a string of the enum value we want.
'
'   Returns: The enum value for the string.

Private Function GetEnumTagIDValue(strTag As String) As enuTagID
    Select Case strTag
        Case "lblProgramRequest":       GetEnumTagIDValue = lblProgramRequest
        Case "chkNewProgram":           GetEnumTagIDValue = chkNewProgram
        Case "lblChkNewProgram":        GetEnumTagIDValue = lblChkNewProgram
        Case "chkModifyExisting":       GetEnumTagIDValue = chkModifyExisting
        Case "lblChkModifyExisting":    GetEnumTagIDValue = lblChkModifyExisting
        Case "lblTxtProgramNumber":     GetEnumTagIDValue = lblTxtProgramNumber
        Case "txtProgramNumber":        GetEnumTagIDValue = txtProgramNumber
        
        Case "lblDtpRequested":         GetEnumTagIDValue = lblDtpRequested
        Case "dtpRequested":            GetEnumTagIDValue = dtpRequested
        Case "lblCmbRequestedBy":       GetEnumTagIDValue = lblCmbRequestedBy
        Case "cmbRequestedBy":          GetEnumTagIDValue = cmbRequestedBy
        Case "lblDtpCompleted":         GetEnumTagIDValue = lblDtpCompleted
        Case "dtpCompleted":            GetEnumTagIDValue = dtpCompleted
        Case "lblCmbCompletedBy":       GetEnumTagIDValue = lblCmbCompletedBy
        Case "cmbCompletedBy":          GetEnumTagIDValue = cmbCompletedBy
        
        Case "txtProgramTitle":         GetEnumTagIDValue = txtProgramTitle
        Case "lblTxtNewProgramNumber":  GetEnumTagIDValue = lblTxtNewProgramNumber
        Case "txtNewProgramNumber":     GetEnumTagIDValue = txtNewProgramNumber
        
        Case "lblTxtOfferingVendor":    GetEnumTagIDValue = lblTxtOfferingVendor
        Case "txtOfferingVendor":       GetEnumTagIDValue = txtOfferingVendor
        Case "chkTrackingOnly":         GetEnumTagIDValue = chkTrackingOnly
        Case "lblChkTrackingOnly":      GetEnumTagIDValue = lblChkTrackingOnly
        Case "chkSales":                GetEnumTagIDValue = chkSales
        Case "lblChkSales":             GetEnumTagIDValue = lblChkSales
        Case "chkPurchasing":           GetEnumTagIDValue = chkPurchasing
        Case "lblChkPurchasing":        GetEnumTagIDValue = lblChkPurchasing
        Case "chkOtherProgramType":     GetEnumTagIDValue = chkOtherProgramType
        Case "lblChkOtherProgramType":  GetEnumTagIDValue = lblChkOtherProgramType
        Case "lblTxtOtherProgramType":  GetEnumTagIDValue = lblTxtOtherProgramType
        Case "txtOtherProgramType":     GetEnumTagIDValue = txtOtherProgramType
        Case "lblLstBuyer":             GetEnumTagIDValue = lblLstBuyer
        Case "lstBuyer":                GetEnumTagIDValue = lstBuyer
        Case "lblLstPayment":           GetEnumTagIDValue = lblLstPayment
        Case "lstPayment":              GetEnumTagIDValue = lstPayment
        Case "lblTxtDescription":       GetEnumTagIDValue = lblTxtDescription
        Case "txtDescription":          GetEnumTagIDValue = txtDescription
        Case "lblDtpStartDate":         GetEnumTagIDValue = lblDtpStartDate
        Case "dtpStartDate":            GetEnumTagIDValue = dtpStartDate
        Case "lblDtpEndDate":           GetEnumTagIDValue = lblDtpEndDate
        Case "dtpEndDate":              GetEnumTagIDValue = dtpEndDate
        Case "chkInvoiced":             GetEnumTagIDValue = chkInvoiced
        Case "lblChkInvoiced":          GetEnumTagIDValue = lblChkInvoiced
        Case "chkReceived":             GetEnumTagIDValue = chkReceived
        Case "lblChkReceived":          GetEnumTagIDValue = lblChkReceived
        Case "chkOrdered":              GetEnumTagIDValue = chkOrdered
        Case "lblChkOrdered":           GetEnumTagIDValue = lblChkOrdered
        Case "lblLstCategory":          GetEnumTagIDValue = lblLstCategory
        Case "lstCategory":             GetEnumTagIDValue = lstCategory
        Case "lblTxtOtherCategory":     GetEnumTagIDValue = lblTxtOtherCategory
        Case "txtOtherCategory":        GetEnumTagIDValue = txtOtherCategory
        
        Case "lblTxtPayingVendor":      GetEnumTagIDValue = lblTxtPayingVendor
        Case "txtPayingVendor":         GetEnumTagIDValue = txtPayingVendor
        Case "chkEmailInvoice":         GetEnumTagIDValue = chkEmailInvoice
        Case "lblChkEmailInvoice":      GetEnumTagIDValue = lblChkEmailInvoice
        Case "lblTxtToEmail1":          GetEnumTagIDValue = lblTxtToEmail1
        Case "txtToEmail1":             GetEnumTagIDValue = txtToEmail1
        Case "lblTxtToEmail2":          GetEnumTagIDValue = lblTxtToEmail2
        Case "txtToEmail2":             GetEnumTagIDValue = txtToEmail2
        Case "lblTxtToEmail3":          GetEnumTagIDValue = lblTxtToEmail3
        Case "txtToEmail3":             GetEnumTagIDValue = txtToEmail3
        Case "lblLstEmailFormat":       GetEnumTagIDValue = lblLstEmailFormat
        Case "lstEmailFormat":          GetEnumTagIDValue = lstEmailFormat
        Case "chkPayBuyingGroup":       GetEnumTagIDValue = chkPayBuyingGroup
        Case "lblChkPayBuyingGroup":    GetEnumTagIDValue = lblChkPayBuyingGroup
        Case "lblTxtBuyingGroup":       GetEnumTagIDValue = lblTxtBuyingGroup
        Case "txtBuyingGroup":          GetEnumTagIDValue = txtBuyingGroup
        Case "lblLstGLAccount":         GetEnumTagIDValue = lblLstGLAccount
        Case "lstGLAccount":            GetEnumTagIDValue = lstGLAccount
        Case "lblTxtMfcProgramID":      GetEnumTagIDValue = lblTxtMfcProgramID
        Case "txtMfcProgramID":         GetEnumTagIDValue = txtMfcProgramID
        
        Case "lblLstRebateType":        GetEnumTagIDValue = lblLstRebateType
        Case "lstRebateType":           GetEnumTagIDValue = lstRebateType
        Case "lblTxtAmount":            GetEnumTagIDValue = lblTxtAmount
        Case "txtAmount":               GetEnumTagIDValue = txtAmount
        Case "lblLstPaidBy":            GetEnumTagIDValue = lblLstPaidBy
        Case "lstPaidBy":               GetEnumTagIDValue = lstPaidBy
        Case "chkRestrictNOI":          GetEnumTagIDValue = chkRestrictNOI
        Case "lblChkRestrictNOI":       GetEnumTagIDValue = lblChkRestrictNOI
        Case "lblTxtNOIAmount":         GetEnumTagIDValue = lblTxtNOIAmount
        Case "txtNOIAmount":            GetEnumTagIDValue = txtNOIAmount
        Case "lblLstNOIUnit":           GetEnumTagIDValue = lblLstNOIUnit
        Case "lstNOIUnit":              GetEnumTagIDValue = lstNOIUnit
    
        Case "lblTxtNotes":             GetEnumTagIDValue = lblTxtNotes
        Case "txtNotes":                GetEnumTagIDValue = txtNotes
                               
        Case "lblRscEligibleItems":     GetEnumTagIDValue = lblRscEligibleItems
        Case "txtLblBtnAddItems":       GetEnumTagIDValue = txtLblBtnAddItems
        Case "txtLblBtnClearItems":     GetEnumTagIDValue = txtLblBtnClearItems
        Case "rscEligibleItems":        GetEnumTagIDValue = rscEligibleItems
        Case "txtItem":                 GetEnumTagIDValue = txtItem
        Case "txtItemAmount":           GetEnumTagIDValue = txtItemAmount
        Case "lstItemPaidBy":           GetEnumTagIDValue = lstItemPaidBy
        
        Case "lblRscEligibleSuppliers": GetEnumTagIDValue = lblRscEligibleSuppliers
        Case "txtLblBtnAddSuppliers":   GetEnumTagIDValue = txtLblBtnAddSuppliers
        Case "txtLblBtnClearSuppliers": GetEnumTagIDValue = txtLblBtnClearSuppliers
        Case "rscEligibleSuppliers":    GetEnumTagIDValue = rscEligibleSuppliers
        Case "txtSupplier":             GetEnumTagIDValue = txtSupplier
        
        Case "lblRscEligibleCustomers": GetEnumTagIDValue = lblRscEligibleCustomers
        Case "txtLblBtnAddCustomers":   GetEnumTagIDValue = txtLblBtnAddCustomers
        Case "txtLblBtnClearCustomers": GetEnumTagIDValue = txtLblBtnClearCustomers
        Case "rscEligibleCustomers":    GetEnumTagIDValue = rscEligibleCustomers
        Case "txtCustomer":             GetEnumTagIDValue = txtCustomer
        Case "chkOnlyNewCustomers":     GetEnumTagIDValue = chkOnlyNewCustomers
        Case "lblChkOnlyNewCustomers":  GetEnumTagIDValue = lblChkOnlyNewCustomers
        Case "lblDtpNewSince":          GetEnumTagIDValue = lblDtpNewSince
        Case "dtpNewSince":             GetEnumTagIDValue = dtpNewSince
    End Select
End Function


' ------------------------------------------------------------------------------
' GetEnumTagIDString - Returns the string of the enum TagID passed.
'
' This function returns the string of the enum Tag ID passed by using a select.
'
'   Parameters;
'       enuTagID: The tag ID of the enum we want the string for.
'
'   Returns: The string of the passed enum.

Private Function GetEnumTagIDString(enuTagID As enuTagID) As String
    Select Case enuTagID
        Case lblProgramRequest:         GetEnumTagIDString = "lblProgramRequest"
        Case chkNewProgram:             GetEnumTagIDString = "chkNewProgram"
        Case lblChkNewProgram:          GetEnumTagIDString = "lblChkNewProgram"
        Case chkModifyExisting:         GetEnumTagIDString = "chkModifyExisting"
        Case lblChkModifyExisting:      GetEnumTagIDString = "lblChkModifyExisting"
        Case lblTxtProgramNumber:       GetEnumTagIDString = "lblTxtProgramNumber"
        Case txtProgramNumber:          GetEnumTagIDString = "txtProgramNumber"

        Case lblDtpRequested:           GetEnumTagIDString = "lblDtpRequested"
        Case dtpRequested:              GetEnumTagIDString = "dtpRequested"
        Case lblCmbRequestedBy:         GetEnumTagIDString = "lblCmbRequestedBy"
        Case cmbRequestedBy:            GetEnumTagIDString = "cmbRequestedBy"
        Case lblDtpCompleted:           GetEnumTagIDString = "lblDtpCompleted"
        Case dtpCompleted:              GetEnumTagIDString = "dtpCompleted"
        Case lblCmbCompletedBy:         GetEnumTagIDString = "lblCmbCompletedBy"
        Case cmbCompletedBy:            GetEnumTagIDString = "cmbCompletedBy"

        Case txtProgramTitle:           GetEnumTagIDString = "txtProgramTitle"
        Case lblTxtNewProgramNumber:    GetEnumTagIDString = "lblTxtNewProgramNumber"
        Case txtNewProgramNumber:       GetEnumTagIDString = "txtNewProgramNumber"

        Case lblTxtOfferingVendor:      GetEnumTagIDString = "lblTxtOfferingVendor"
        Case txtOfferingVendor:         GetEnumTagIDString = "txtOfferingVendor"
        Case chkTrackingOnly:           GetEnumTagIDString = "chkTrackingOnly"
        Case lblChkTrackingOnly:        GetEnumTagIDString = "lblChkTrackingOnly"
        Case chkSales:                  GetEnumTagIDString = "chkSales"
        Case lblChkSales:               GetEnumTagIDString = "lblChkSales"
        Case chkPurchasing:             GetEnumTagIDString = "chkPurchasing"
        Case lblChkPurchasing:          GetEnumTagIDString = "lblChkPurchasing"
        Case chkOtherProgramType:       GetEnumTagIDString = "chkOtherProgramType"
        Case lblChkOtherProgramType:    GetEnumTagIDString = "lblChkOtherProgramType"
        Case lblTxtOtherProgramType:    GetEnumTagIDString = "lblTxtOtherProgramType"
        Case txtOtherProgramType:       GetEnumTagIDString = "txtOtherProgramType"
        Case lblLstBuyer:               GetEnumTagIDString = "lblLstBuyer"
        Case lstBuyer:                  GetEnumTagIDString = "lstBuyer"
        Case lblLstPayment:             GetEnumTagIDString = "lblLstPayment"
        Case lstPayment:                GetEnumTagIDString = "lstPayment"
        Case lblTxtDescription:         GetEnumTagIDString = "lblTxtDescription"
        Case txtDescription:            GetEnumTagIDString = "txtDescription"
        Case lblDtpStartDate:           GetEnumTagIDString = "lblDtpStartDate"
        Case dtpStartDate:              GetEnumTagIDString = "dtpStartDate"
        Case lblDtpEndDate:             GetEnumTagIDString = "lblDtpEndDate"
        Case dtpEndDate:                GetEnumTagIDString = "dtpEndDate"
        Case chkInvoiced:               GetEnumTagIDString = "chkInvoiced"
        Case lblChkInvoiced:            GetEnumTagIDString = "lblChkInvoiced"
        Case chkReceived:               GetEnumTagIDString = "chkReceived"
        Case lblChkReceived:            GetEnumTagIDString = "lblChkReceived"
        Case chkOrdered:                GetEnumTagIDString = "chkOrdered"
        Case lblChkOrdered:             GetEnumTagIDString = "lblChkOrdered"
        Case lblLstCategory:            GetEnumTagIDString = "lblLstCategory"
        Case lstCategory:               GetEnumTagIDString = "lstCategory"
        Case lblTxtOtherCategory:       GetEnumTagIDString = "lblTxtOtherCategory"
        Case txtOtherCategory:          GetEnumTagIDString = "txtOtherCategory"

        Case lblTxtPayingVendor:        GetEnumTagIDString = "lblTxtPayingVendor"
        Case txtPayingVendor:           GetEnumTagIDString = "txtPayingVendor"
        Case chkEmailInvoice:           GetEnumTagIDString = "chkEmailInvoice"
        Case lblChkEmailInvoice:        GetEnumTagIDString = "lblChkEmailInvoice"
        Case lblTxtToEmail1:            GetEnumTagIDString = "lblTxtToEmail1"
        Case txtToEmail1:               GetEnumTagIDString = "txtToEmail1"
        Case lblTxtToEmail2:            GetEnumTagIDString = "lblTxtToEmail2"
        Case txtToEmail2:               GetEnumTagIDString = "txtToEmail2"
        Case lblTxtToEmail3:            GetEnumTagIDString = "lblTxtToEmail3"
        Case txtToEmail3:               GetEnumTagIDString = "txtToEmail3"
        Case lblLstEmailFormat:         GetEnumTagIDString = "lblLstEmailFormat"
        Case lstEmailFormat:            GetEnumTagIDString = "lstEmailFormat"
        Case chkPayBuyingGroup:         GetEnumTagIDString = "chkPayBuyingGroup"
        Case lblChkPayBuyingGroup:      GetEnumTagIDString = "lblChkPayBuyingGroup"
        Case lblTxtBuyingGroup:         GetEnumTagIDString = "lblTxtBuyingGroup"
        Case txtBuyingGroup:            GetEnumTagIDString = "txtBuyingGroup"
        Case lblLstGLAccount:           GetEnumTagIDString = "lblLstGLAccount"
        Case lstGLAccount:              GetEnumTagIDString = "lstGLAccount"
        Case lblTxtMfcProgramID:        GetEnumTagIDString = "lblTxtMfcProgramID"
        Case txtMfcProgramID:           GetEnumTagIDString = "txtMfcProgramID"

        Case lblLstRebateType:          GetEnumTagIDString = "lblLstRebateType"
        Case lstRebateType:             GetEnumTagIDString = "lstRebateType"
        Case lblTxtAmount:              GetEnumTagIDString = "lblTxtAmount"
        Case txtAmount:                 GetEnumTagIDString = "txtAmount"
        Case lblLstPaidBy:              GetEnumTagIDString = "lblLstPaidBy"
        Case lstPaidBy:                 GetEnumTagIDString = "lstPaidBy"
        Case chkRestrictNOI:            GetEnumTagIDString = "chkRestrictNOI"
        Case lblChkRestrictNOI:         GetEnumTagIDString = "lblChkRestrictNOI"
        Case lblTxtNOIAmount:           GetEnumTagIDString = "lblTxtNOIAmount"
        Case txtNOIAmount:              GetEnumTagIDString = "txtNOIAmount"
        Case lblLstNOIUnit:             GetEnumTagIDString = "lblLstNOIUnit"
        Case lstNOIUnit:                GetEnumTagIDString = "lstNOIUnit"

        Case lblTxtNotes:               GetEnumTagIDString = "lblTxtNotes"
        Case txtNotes:                  GetEnumTagIDString = "txtNotes"

        Case lblRscEligibleItems:       GetEnumTagIDString = "lblRscEligibleItems"
        Case txtLblBtnAddItems:         GetEnumTagIDString = "txtLblBtnAddItems"
        Case txtLblBtnClearItems:       GetEnumTagIDString = "txtLblBtnClearItems"
        Case rscEligibleItems:          GetEnumTagIDString = "rscEligibleItems"
        Case txtItem:                   GetEnumTagIDString = "txtItem"
        Case txtItemAmount:             GetEnumTagIDString = "txtItemAmount"
        Case lstItemPaidBy:             GetEnumTagIDString = "lstItemPaidBy"

        Case lblRscEligibleSuppliers:   GetEnumTagIDString = "lblRscEligibleSuppliers"
        Case txtLblBtnAddSuppliers:     GetEnumTagIDString = "txtLblBtnAddSuppliers"
        Case txtLblBtnClearSuppliers:   GetEnumTagIDString = "txtLblBtnClearSuppliers"
        Case rscEligibleSuppliers:      GetEnumTagIDString = "rscEligibleSuppliers"
        Case txtSupplier:               GetEnumTagIDString = "txtSupplier"

        Case lblRscEligibleCustomers:   GetEnumTagIDString = "lblRscEligibleCustomers"
        Case txtLblBtnAddCustomers:     GetEnumTagIDString = "txtLblBtnAddCustomers"
        Case txtLblBtnClearCustomers:   GetEnumTagIDString = "txtLblBtnClearCustomers"
        Case rscEligibleCustomers:      GetEnumTagIDString = "rscEligibleCustomers"
        Case txtCustomer:               GetEnumTagIDString = "txtCustomer"
        Case chkOnlyNewCustomers:       GetEnumTagIDString = "chkOnlyNewCustomers"
        Case lblChkOnlyNewCustomers:    GetEnumTagIDString = "lblChkOnlyNewCustomers"
        Case lblDtpNewSince:            GetEnumTagIDString = "lblDtpNewSince"
        Case dtpNewSince:               GetEnumTagIDString = "dtpNewSince"
    End Select
End Function


' ------------------------------------------------------------------------------
' IsAlphaNumeric - Check to see if a statement is alphanumeric or not.
'
' This function checks the passed statement letter by letter to make sure they
' are all number or letters only. It may or may not count spaces as letters
' based on users preference. It returns true only if all charters were good.
'
'   Parameters;
'       strStatment:    The statement we are checking.
'       blnAllowSpaces: If spaces should count as alphanumeric or not.
'
'   Returns: True if all charters were alphanumeric, False otherwise.

Private Function IsAlphaNumeric(strStatment As String, Optional blnAllowSpaces As Boolean = False) As Boolean
    Dim intPos As Integer
    
    For intPos = 1 To Len(strStatment)
        Select Case Asc(Mid(strStatment, intPos, 1))
            
            'ASCII values for 0 to 9, lowercase, and uppercase letters.
            Case 48 To 57, 65 To 90, 97 To 122
            
            'ASCII value for space.
            Case 32
                If Not blnAllowSpaces Then
                    IsAlphaNumeric = False
                    Exit Function
                End If
                
            Case Else
                IsAlphaNumeric = False
                Exit Function

        End Select
    Next intPos
    
    IsAlphaNumeric = True
End Function


' ------------------------------------------------------------------------------
' IsPlaceholder - Checks to see if a control is set to its placeholder text.
'
' This function checks a user input field against the placeholder text for
' that fields tag id in the placeholder string array. If it is a match, then it
' returns true, if it dose not find a match it returns false. We first check to
' see if our default placeholder is blank. If it is, then we need to call our
' set placeholders function to set the default values before we compare. We need
' to do this as users improperly closing the form can cause these values to
' get reset to null.
'
'   Parameters;
'       occTarget:  The user input control that needs checked.
'
'   Returns: True if text is same as placeholder, otherwise False.

Private Function IsPlaceholder(occTarget As ContentControl) As Boolean
    Dim enuID As enuTagID
    Dim strDefault As String
    enuID = GetEnumTagIDValue(occTarget.Tag)
    strDefault = strPlaceholder(enuID)
    
    If strDefault = "" Then
        InitializePlaceholders
        strDefault = strPlaceholder(enuID)
    End If

    If occTarget.Range.Text = strPlaceholder(enuID) Then
        IsPlaceholder = True
        Exit Function
    End If
    
    IsPlaceholder = False
End Function


' ------------------------------------------------------------------------------
' IsPrintSpace - Checks to see if a control is set to its print spaces.
'
' This function checks a user input field against the print spaces text for
' that fields tag id in the print space string array. If it is a match, then it
' returns true, if it dose not find a match it returns false. We first check to
' see if our default sprint space is blank. If it is, then we need to call our
' set print spaces function to set the default values before we compare. We need
' to do this as users improperly closing the form can cause these values to
' get reset to null.
'
'   Parameters;
'       occTarget:  The user input control that needs checked.
'
'   Returns: True if text is same as print space, otherwise False.

Private Function IsPrintSpace(occTarget As ContentControl) As Boolean
    Dim enuID As enuTagID
    Dim strSpaces As String
    enuID = GetEnumTagIDValue(occTarget.Tag)
    strSpaces = strPrintSpace(enuID)
    
    If strSpaces = "" Then
        InitializePrintSpace
        strSpaces = strPrintSpace(enuID)
    End If

    If occTarget.Range.Text = strPrintSpace(enuID) Then
        IsPrintSpace = True
        Exit Function
    End If
    
    IsPrintSpace = False
End Function


' ------------------------------------------------------------------------------
' ValidateAmount - Checks a $ or % ranges value to ensure good data and format.
'
' This function checks the user input of the passed control, stripping out any
' existing formating, ensuring it is numeric, and giving it the proper format.
' If there is a problem with the data it will generate a message box with
' details on how to correct, and set the users focus on the range. It returns a
' boolean to tell if the checked input was valid or not. It checks the value
' against the default placeholder values. If it is one of them it will ignore
' the target and pass validation.
'
'   Parameters;
'       occTarget:  The user input control that needs validated.
'       strFormat:  Which way to format range's text.
'
'   Returns: True if data passed validation, False if an error was thrown.

Private Function ValidateAmount(occTarget As ContentControl, Optional enuFormat As enuAmountFormat = RationalNumber) As Boolean
    Dim strAmount As String
    strAmount = occTarget.Range.Text
        
    If Not IsPlaceholder(occTarget) And strAmount <> "Disabled" And strAmount <> " " Then
        
        If Left(strAmount, 1) = "$" Then
            strAmount = Right(strAmount, Len(strAmount) - 1)
        ElseIf Right(strAmount, 1) = "%" Then
            strAmount = Left(strAmount, Len(strAmount) - 1)
        End If
        
        If Not IsNumeric(strAmount) Then
            Call MsgBox("Please enter a numeric amount.", vbOKOnly, "Validation Error!")
            occTarget.Range.Select
            ValidateAmount = False
            Exit Function
        End If
        
        If enuFormat = Cash Then
            strAmount = Format(strAmount, "0.00")
            occTarget.Range.Text = "$" & strAmount
            
        ElseIf enuFormat = Percentage Then
            strAmount = Format(strAmount, "0.000")
            occTarget.Range.Text = strAmount & "%"
            
        ElseIf enuFormat = WholeNumber Then
            strAmount = Format(strAmount, "0")
            occTarget.Range.Text = strAmount
            
        ElseIf enuFormat = RationalNumber Then
            strAmount = Format(strAmount, "0.###")
            If Right(strAmount, 1) = "." Then: strAmount = Left(strAmount, Len(strAmount) - 1)
            occTarget.Range.Text = strAmount
        End If
        
    End If
    
    ValidateAmount = True
End Function


' ------------------------------------------------------------------------------
' AutofillCmbRequestedBy - Fills in the requester from the buyer.
'
' This subroutine checks to see if the requester is blank, and a buyer has
' been filled in. If so it copies the buyer to the requester.

Private Sub AutofillCmbRequestedBy()
    Dim occBuyer As ContentControl
    Dim occRequester As ContentControl
    
    Set occBuyer = ActiveDocument.SelectContentControlsByTag("lstBuyer")(1)
    Set occRequester = ActiveDocument.SelectContentControlsByTag("cmbRequestedBy")(1)
    
    If IsPlaceholder(occRequester) And Not IsPlaceholder(occBuyer) Then: occRequester.Range.Text = occBuyer.Range.Text
End Sub


' ------------------------------------------------------------------------------
' AutofillLstBuyer - Fills in the buyer from the requester if it is a buyer.
'
' This subroutine checks to see if the buyer is blank and the requester has been
' filled in. It then check to see if the requester is in the list of buyers. If
' so it selects that buyer, saving our user a bit of work.

Private Sub AutofillLstBuyer()
    Dim occBuyer As ContentControl
    Dim oleBuyer As ContentControlListEntry
    Dim occRequester As ContentControl
    
    Set occBuyer = ActiveDocument.SelectContentControlsByTag("lstBuyer")(1)
    Set occRequester = ActiveDocument.SelectContentControlsByTag("cmbRequestedBy")(1)
    
    If IsPlaceholder(occBuyer) And Not IsPlaceholder(occRequester) Then
        For Each oleBuyer In occBuyer.DropdownListEntries
            If occRequester.Range.Text = oleBuyer.Text Then
                occBuyer.DropdownListEntries(oleBuyer.Index).Select
                Exit For
            End If
        Next oleBuyer
    End If
End Sub


' ------------------------------------------------------------------------------
' AutofillTxtPayingVendor - Fills in the paying vendor from offering vendor.
'
' This subroutine checks to see if the paying vendor is blank and an offering
' vendor has been filled in. If so it copies the offering vendor over to the
' paying vendor field to save the user some work.

Private Sub AutofillTxtPayingVendor()
    Dim occOfferingVendor As ContentControl
    Dim occPayingVendor As ContentControl
    
    Set occOfferingVendor = ActiveDocument.SelectContentControlsByTag("txtOfferingVendor")(1)
    Set occPayingVendor = ActiveDocument.SelectContentControlsByTag("txtPayingVendor")(1)
    
    If IsPlaceholder(occPayingVendor) And Not (IsPlaceholder(occOfferingVendor) Or occOfferingVendor.Range.Text = "Disabled") _
    Then: occPayingVendor.Range.Text = occOfferingVendor.Range.Text
End Sub


' ------------------------------------------------------------------------------
' AutofillTxtItemAmount - Fills in a default payment and amount for an item.
'
' This subroutine finds the row that the passed item is on, and retrieves the
' corresponding amount and payment. If the items placeholder text has been
' overwritten, and the user has not yet selected a payment type, and the user
' set up a default payment type if will fill in the payment type based off the
' default. It will then check to see if the amount was left blank, there is a
' default amount set up, and the payment method matches the default amount. If
' so it will also fill in the amount with the default amount.
'
'   Parameters;
'       occItem: The content control whose user input we want checked.

Private Sub AutofillTxtItemAmount(occItem As ContentControl)
    Dim tblSection As Table
    Dim intRow As Integer
    Dim intAccessPayment As Integer
    Dim occAmount As ContentControl
    Dim occPayment As ContentControl
    Dim occDefaultAmount As ContentControl
    Dim occDefaultPayment As ContentControl
   
    Set tblSection = occItem.ParentContentControl.RepeatingSectionItems.Item(1).Range.Tables(1)
    intRow = GetRow(occItem.ID, occItem.ParentContentControl.Tag)
    intAccessPayment = AccessIndex(intRow, 3, tblSection.Rows.Count, 3)
    Set occAmount = tblSection.Rows(intRow).Cells(2).Range.ContentControls(1)
    Set occPayment = tblSection.Rows(intRow).Cells(3).Range.ContentControls(intAccessPayment)
    Set occDefaultAmount = ActiveDocument.SelectContentControlsByTag("txtAmount")(1)
    Set occDefaultPayment = ActiveDocument.SelectContentControlsByTag("lstPaidBy")(1)
    
    'Check to make sure they have an item filled in, have not selected a payment, and the default is not blank or disabled.
    If Not IsPlaceholder(occItem) _
    And IsPlaceholder(occPayment) _
    And Not IsPlaceholder(occDefaultPayment) _
    And occDefaultPayment.Range.Text <> "Disabled" Then
            occPayment.Type = wdContentControlText
            occPayment.Range.Text = occDefaultPayment.Range.Text
            occPayment.Type = wdContentControlDropdownList
    End If
    
    'Checking to make sure they have not filled in an amount, the payment type matches the default, and is not blank.
    If IsPlaceholder(occAmount) _
    And Not IsPlaceholder(occPayment) _
    And occPayment.Range.Text = occDefaultPayment.Range.Text Then
            occAmount.Range.Text = occDefaultAmount.Range.Text
    End If
End Sub


' ------------------------------------------------------------------------------
' CheckTxtItem - Checks user enter data to make sure it is valid.
'
' This function checks the user entered data in one of the txtItem fields. It
' verifies that the user has overwritten the placeholder text. Then converts the
' input to uppercase, and check to make sure it is less than 10 charters,
' contains only alphanumeric data with no spaces. And finally checks it against
' all other items on the list to make sure it is not a duplicate. If it passes
' all tests then the function returns true to show that it was good input, but
' if it fails at any point the function will give the user a description of what
' they did wrong, highlight the field, return false and exit the function.
'
'   Parameters;
'       occItem: The content control whose user input we want checked.
'
'   Returns: Boolean showing if user input passed validation or not.

Private Function CheckTxtItem(occItem As ContentControl) As Boolean
    Dim ocsItems As ContentControls
    Dim strItem As String
    Dim intIndex As Integer
    
    
    If Not IsPlaceholder(occItem) Then
        strItem = UCase(occItem.Range.Text)
        occItem.Range.Text = strItem
        
        If Len(strItem) > 10 Then
            Call MsgBox("Please enter 10 digits or less.", vbOKOnly, "Validation Error!")
            occItem.Range.Select
            CheckTxtItem = False
            Exit Function
        End If
        
        If Not IsAlphaNumeric(strItem) Then
            Call MsgBox("Please enter only alpha-numeric characters." & vbNewLine & "Spaces are not permitted.", vbOKOnly, "Validation Error!")
            occItem.Range.Select
            CheckTxtItem = False
            Exit Function
        End If
        
        Set ocsItems = ActiveDocument.SelectContentControlsByTag(occItem.Tag)
        For intIndex = 1 To ocsItems.Count
            If occItem.ID <> ocsItems(intIndex).ID And strItem = ocsItems(intIndex).Range.Text Then
                Call MsgBox("There is all ready an item with this number." & vbNewLine & "Duplicate entries are not allowed.", vbOKOnly, "Validation Error!")
                occItem.Range.Select
                CheckTxtItem = False
                Exit Function
            End If
        Next intIndex
    End If
    
    CheckTxtItem = True
End Function


' ------------------------------------------------------------------------------
' CheckTxtItemAmount - Checks user enter data to make sure it is valid.
'
' This function checks the user entered data in one of the txtItemAmount fields.
' It uses the id of the passed content control to find out what row of the RSC
' it is on. It uses that row information to get the payment type for the same
' row. Then it calls the ValidateAmount function to evaluate the user input and
' returns a boolean showing if we had valid data or not.
'
'   Parameters;
'       occItemAmount: The content control whose user input we want checked.
'
'   Returns: Boolean showing if user input passed validation or not.

Private Function CheckTxtItemAmount(occItemAmount As ContentControl) As Boolean
    Dim intRow As Integer
    Dim tblSection As Table
    Dim intAccessPayment As Integer
    Dim enuPayment As enuAmountFormat
    
    intRow = GetRow(occItemAmount.ID, occItemAmount.ParentContentControl.Tag)
    
    If intRow = 0 Then
        CheckTxtItemAmount = False
        Exit Function
    End If
    
    Set tblSection = occItemAmount.ParentContentControl.RepeatingSectionItems.Item(1).Range.Tables(1)
    intAccessPayment = AccessIndex(intRow, 3, tblSection.Rows.Count, 3)
    enuPayment = GetEnumAmountFormatValue(Left(tblSection.Rows(intRow).Cells(3).Range.ContentControls(intAccessPayment).Range.Text, 1))
    
    If Not ValidateAmount(occItemAmount, enuPayment) Then
        CheckTxtItemAmount = False
        Exit Function
    End If

    CheckTxtItemAmount = True
End Function


' ------------------------------------------------------------------------------
' CheckTxtSupplier - Checks user enter data to make sure it is valid.
'
' This function checks the user entered data in one of the txtSupplier fields.
' It verifies that the user has overwritten the placeholder text. Then check to
' only whole numbers were entered. Last it checks the user data against the
' other suppliers on the list to make sure it is not a duplicate. If it passes
' all tests then the function returns true to show that it was good input, but
' if it fails at any point the function will give the user a description of what
' they did wrong, highlight the field, return false and exit the function.
'
'   Parameters;
'       occItem: The content control whose user input we want checked.
'
'   Returns: Boolean showing if user input passed validation or not.

Private Function CheckTxtSupplier(occSupplier As ContentControl) As Boolean
    Dim ocsSuppliers As ContentControls
    Dim strSupplier As String
    Dim intIndex As Integer
    
    If Not IsPlaceholder(occSupplier) Then
        strSupplier = occSupplier.Range.Text
        
        If Not ValidateAmount(occSupplier, WholeNumber) Then
            CheckTxtSupplier = False
            Exit Function
        End If
        
        Set ocsSuppliers = ActiveDocument.SelectContentControlsByTag(occSupplier.Tag)
        For intIndex = 1 To ocsSuppliers.Count
            If occSupplier.ID <> ocsSuppliers(intIndex).ID And strSupplier = ocsSuppliers(intIndex).Range.Text Then
                Call MsgBox("There is all ready an supplier with this number." & vbNewLine & "Duplicate entries are not allowed.", vbOKOnly, "Validation Error!")
                occSupplier.Range.Select
                CheckTxtSupplier = False
                Exit Function
            End If
        Next intIndex
    End If
    
    CheckTxtSupplier = True
End Function


' ------------------------------------------------------------------------------
' CheckTxtCustomer - Checks user enter data to make sure it is valid.
'
' This function checks the user entered data in one of the txtCustomer fields.
' It verifies that the user has overwritten the placeholder text. Then check to
' only whole numbers were entered. Last it checks the user data against the
' other customers on the list to make sure it is not a duplicate. If it passes
' all tests then the function returns true to show that it was good input, but
' if it fails at any point the function will give the user a description of what
' they did wrong, highlight the field, return false and exit the function.
'
'   Parameters;
'       occItem: The content control whose user input we want checked.
'
'   Returns: Boolean showing if user input passed validation or not.

Private Function CheckTxtCustomer(occCustomer As ContentControl) As Boolean
    Dim ocsCustomers As ContentControls
    Dim strCustomer As String
    Dim intIndex As Integer
    
    If Not IsPlaceholder(occCustomer) Then
        strCustomer = occCustomer.Range.Text
        
        If Not ValidateAmount(occCustomer, WholeNumber) Then
            CheckTxtCustomer = False
            Exit Function
        End If
        
        Set ocsCustomers = ActiveDocument.SelectContentControlsByTag(occCustomer.Tag)
        For intIndex = 1 To ocsCustomers.Count
            If occCustomer.ID <> ocsCustomers(intIndex).ID And strCustomer = ocsCustomers(intIndex).Range.Text Then
                Call MsgBox("There is all ready an customer with this number." & vbNewLine & "Duplicate entries are not allowed.", vbOKOnly, "Validation Error!")
                occCustomer.Range.Select
                CheckTxtCustomer = False
                Exit Function
            End If
        Next intIndex
    End If
    
    CheckTxtCustomer = True
End Function


' ------------------------------------------------------------------------------
' ToggelChkNewProgram - Turns on or off features based on the New Program field.
'
' This will check the sales and modify program check boxes and turn other form
' fields on or off based on their states. Having new program unchecked will
' will enable the program number field, and check it will disable that field.
' Leaving existing unchecked will enable the offering vendor field, and it
' will be disabled if existing is checked. Finally we select the next field.

Private Sub ToggelChkNewProgram()
    Dim occChkNewProgram As ContentControl
    Dim occChkModifyExisting As ContentControl
    Dim occLblTxtProgramNumber As ContentControl
    Dim occTxtProgramNumber As ContentControl
    Dim occLblTxtOfferingVendor As ContentControl
    Dim occTxtOfferingVendor As ContentControl
    
    Set occChkNewProgram = ActiveDocument.SelectContentControlsByTag("chkNewProgram")(1)
    Set occChkModifyExisting = ActiveDocument.SelectContentControlsByTag("chkModifyExisting")(1)
    Set occLblTxtProgramNumber = ActiveDocument.SelectContentControlsByTag("lblTxtProgramNumber")(1)
    Set occTxtProgramNumber = ActiveDocument.SelectContentControlsByTag("txtProgramNumber")(1)
    Set occLblTxtOfferingVendor = ActiveDocument.SelectContentControlsByTag("lblTxtOfferingVendor")(1)
    Set occTxtOfferingVendor = ActiveDocument.SelectContentControlsByTag("txtOfferingVendor")(1)

    If Not occChkNewProgram.Checked Then
        occLblTxtProgramNumber.DefaultTextStyle = "Heading 4"
    
        occTxtProgramNumber.LockContents = False
        If occTxtProgramNumber.Range.Text = "Disabled" Then: occTxtProgramNumber.Range.Text = ""
        occTxtProgramNumber.DefaultTextStyle = "ContentControlsSmall"
        occTxtProgramNumber.Appearance = wdContentControlBoundingBox
    Else
        occLblTxtProgramNumber.DefaultTextStyle = "OffSmall"
    
        occTxtProgramNumber.LockContents = False
        occTxtProgramNumber.Range.Text = "Disabled"
        occTxtProgramNumber.DefaultTextStyle = "OffSmall"
        occTxtProgramNumber.Appearance = wdContentControlHidden
        occTxtProgramNumber.LockContents = True
    End If
    
    If Not occChkModifyExisting.Checked Then
        occLblTxtOfferingVendor.DefaultTextStyle = "Heading 3"
    
        occTxtOfferingVendor.LockContents = False
        If occTxtOfferingVendor.Range.Text = "Disabled" Then: occTxtOfferingVendor.Range.Text = ""
        occTxtOfferingVendor.DefaultTextStyle = "ContentControls1"
        occTxtOfferingVendor.Appearance = wdContentControlBoundingBox
    Else
        occLblTxtOfferingVendor.DefaultTextStyle = "OffSmall"
    
        occTxtOfferingVendor.LockContents = False
        occTxtOfferingVendor.Range.Text = "Disabled"
        occTxtOfferingVendor.DefaultTextStyle = "OffSmall"
        occTxtOfferingVendor.Appearance = wdContentControlHidden
        occTxtOfferingVendor.LockContents = True
    End If
    
    If occChkModifyExisting.Checked Then
        ActiveDocument.SelectContentControlsByTag("txtProgramNumber")(1).Range.Select
    Else
        ActiveDocument.SelectContentControlsByTag("txtOfferingVendor")(1).Range.Select
    End If
End Sub


' ------------------------------------------------------------------------------
' ToggelTxtNewProgramNumber - Turns on or off features based on being completed.
'
' This will check to see if a completed date and name has been filled in. If so
' then it will display the field to enter a new program number, or turn it off
' if it has not been completed.

Private Sub ToggelTxtNewProgramNumber()
    Dim occCmbCompletedBy As ContentControl
    Dim occDtpCompleted As ContentControl
    Dim occLblTxtNewProgramNumber As ContentControl
    Dim occTxtNewProgramNumber As ContentControl
    
    Set occCmbCompletedBy = ActiveDocument.SelectContentControlsByTag("cmbCompletedBy")(1)
    Set occDtpCompleted = ActiveDocument.SelectContentControlsByTag("dtpCompleted")(1)
    Set occLblTxtNewProgramNumber = ActiveDocument.SelectContentControlsByTag("lblTxtNewProgramNumber")(1)
    Set occTxtNewProgramNumber = ActiveDocument.SelectContentControlsByTag("txtNewProgramNumber")(1)

    If Not IsPlaceholder(occCmbCompletedBy) And Not IsPlaceholder(occDtpCompleted) Then
        occLblTxtNewProgramNumber.LockContents = False
        occLblTxtNewProgramNumber.Range.Text = vbTab & "Program #: "
        occLblTxtNewProgramNumber.LockContents = True
        
        occTxtNewProgramNumber.LockContents = False
        occTxtNewProgramNumber.Range.Text = ""
        occTxtNewProgramNumber.DefaultTextStyle = "NewProgramNumber"
    Else
        occLblTxtNewProgramNumber.LockContents = False
        occLblTxtNewProgramNumber.Range.Text = ""
        occLblTxtNewProgramNumber.LockContents = True
        
        occTxtNewProgramNumber.LockContents = False
        occTxtNewProgramNumber.Range.Text = " "
        occTxtNewProgramNumber.DefaultTextStyle = "Off"
        occTxtNewProgramNumber.Appearance = wdContentControlHidden
        occTxtNewProgramNumber.LockContents = True
    End If
End Sub


' ------------------------------------------------------------------------------
' ToggelChkTrackingOnly - Turns on or off features based on chkTrackingOnly.
'
' This will check to see chkTrackingOnly is checked. If so it will set the field
' for payment and category to tracking only and disable user editing of them. If
' not it will reset those fields. As well it call the toggle for category so
' the if other field get a proper state set. Finally it calls update title, as
' changes to the payment and category fields can effect the title.

Private Sub ToggelChkTrackingOnly()
    Dim occChkTrackingOnly As ContentControl
    Dim occLstPayment As ContentControl
    Dim occLstCategory As ContentControl
    Dim occLblTxtOtherCategory As ContentControl
    
    Set occChkTrackingOnly = ActiveDocument.SelectContentControlsByTag("chkTrackingOnly")(1)
    Set occLstPayment = ActiveDocument.SelectContentControlsByTag("lstPayment")(1)
    Set occLstCategory = ActiveDocument.SelectContentControlsByTag("lstCategory")(1)
    Set occLblTxtOtherCategory = ActiveDocument.SelectContentControlsByTag("lblTxtOtherCategory")(1)

    If occChkTrackingOnly.Checked Then
        occLstPayment.LockContents = False
        occLstPayment.DropdownListEntries(5).Select
        occLstPayment.DefaultTextStyle = "Locked"
        occLstPayment.Appearance = wdContentControlHidden
        occLstPayment.LockContents = True
        
        occLstCategory.LockContents = False
        occLstCategory.DefaultTextStyle = "Locked"
        occLstCategory.DropdownListEntries(7).Select
        occLstCategory.Appearance = wdContentControlHidden
        occLstCategory.LockContents = True
    Else
        occLstPayment.LockContents = False
        Call ResetList(occLstPayment.Tag, , True)
        occLstPayment.DefaultTextStyle = "ContentControls1"
        occLstPayment.Appearance = wdContentControlBoundingBox
        
        occLstCategory.LockContents = False
        Call ResetList(occLstCategory.Tag, , True)
        occLstCategory.DefaultTextStyle = "ContentControls1"
        occLstCategory.Appearance = wdContentControlBoundingBox
    End If
    
    ToggelLstCategory
    UpdateTitle
    
    If ActiveDocument.SelectContentControlsByTag("chkSales")(1).Checked Or ActiveDocument.SelectContentControlsByTag("chkPurchasing")(1).Checked Then
        ActiveDocument.SelectContentControlsByTag("lstBuyer")(1).Range.Select
    Else
        ActiveDocument.SelectContentControlsByTag("txtOtherProgramType")(1).Range.Select
    End If
End Sub


' ------------------------------------------------------------------------------
' ToggelChkSales - Turns on or off features based on sales and other fields.
'
' This will see if the user has check the sales box. If so it turns off the
' options to base the program on invoice, received, or ordered. Makes sure that
' the NOI options are enabled, as well as the RSC for customers and only new.
' if sales is not checked it will turn on the invoice options, and off the
' NOI options. If neither purchasing or other were selected it will allow
' the customers RSC and only new fields to stay on, otherwise it turns them off.
' Finally we call addition toggles that need to follow this one.

Private Sub ToggelChkSales()
    Dim occChkSales As ContentControl
    Dim occChkPurchasing As ContentControl
    Dim occChkOtherProgramType As ContentControl
    Dim occChkInvoiced As ContentControl
    Dim occLblChkInvoiced As ContentControl
    Dim occChkReceived As ContentControl
    Dim occLblChkReceived As ContentControl
    Dim occChkOrdered As ContentControl
    Dim occLblChkOrdered As ContentControl
    Dim occChkRestrictNOI As ContentControl
    Dim occLblChkRestrictNOI As ContentControl
    Dim occLblRscEligibleCustomers As ContentControl
    Dim occTxtLblBtnAddCustomers As ContentControl
    Dim occTxtLblBtnClearCustomers As ContentControl
    Dim occRscEligibleCustomers As ContentControl
    Dim occChkOnlyNewCustomers As ContentControl
    Dim occLblChkOnlyNewCustomers As ContentControl
    Dim ocsRSCItems As RepeatingSectionItemColl
    Dim ocsCustomers As ContentControls
    Dim occCustomer As ContentControl
    Dim intIndex As Integer
    
    Set occChkSales = ActiveDocument.SelectContentControlsByTag("chkSales")(1)
    Set occChkPurchasing = ActiveDocument.SelectContentControlsByTag("chkPurchasing")(1)
    Set occChkOtherProgramType = ActiveDocument.SelectContentControlsByTag("chkOtherProgramType")(1)
    Set occChkInvoiced = ActiveDocument.SelectContentControlsByTag("chkInvoiced")(1)
    Set occLblChkInvoiced = ActiveDocument.SelectContentControlsByTag("lblChkInvoiced")(1)
    Set occChkReceived = ActiveDocument.SelectContentControlsByTag("chkReceived")(1)
    Set occLblChkReceived = ActiveDocument.SelectContentControlsByTag("lblChkReceived")(1)
    Set occChkOrdered = ActiveDocument.SelectContentControlsByTag("chkOrdered")(1)
    Set occLblChkOrdered = ActiveDocument.SelectContentControlsByTag("lblChkOrdered")(1)
    Set occChkRestrictNOI = ActiveDocument.SelectContentControlsByTag("chkRestrictNOI")(1)
    Set occLblChkRestrictNOI = ActiveDocument.SelectContentControlsByTag("lblChkRestrictNOI")(1)
    Set occLblRscEligibleCustomers = ActiveDocument.SelectContentControlsByTag("lblRscEligibleCustomers")(1)
    Set occTxtLblBtnAddCustomers = ActiveDocument.SelectContentControlsByTag("txtLblBtnAddCustomers")(1)
    Set occTxtLblBtnClearCustomers = ActiveDocument.SelectContentControlsByTag("txtLblBtnClearCustomers")(1)
    Set occRscEligibleCustomers = ActiveDocument.SelectContentControlsByTag("rscEligibleCustomers")(1)
    Set occChkOnlyNewCustomers = ActiveDocument.SelectContentControlsByTag("chkOnlyNewCustomers")(1)
    Set occLblChkOnlyNewCustomers = ActiveDocument.SelectContentControlsByTag("lblChkOnlyNewCustomers")(1)
    Set ocsRSCItems = ActiveDocument.SelectContentControlsByTag("rscEligibleCustomers")(1).RepeatingSectionItems
    Set ocsCustomers = ActiveDocument.SelectContentControlsByTag("txtCustomer")
        
    If occChkSales.Checked Then
        occChkInvoiced.LockContents = False
        occChkInvoiced.Checked = False
        occChkInvoiced.DefaultTextStyle = "Off"
        occChkInvoiced.Appearance = wdContentControlHidden
        occChkInvoiced.LockContents = True
        
        occLblChkInvoiced.DefaultTextStyle = "Off"
        
        occChkReceived.LockContents = False
        occChkReceived.Checked = False
        occChkReceived.DefaultTextStyle = "Off"
        occChkReceived.Appearance = wdContentControlHidden
        occChkReceived.LockContents = True
        
        occLblChkReceived.DefaultTextStyle = "Off"
        
        occChkOrdered.LockContents = False
        occChkOrdered.Checked = False
        occChkOrdered.DefaultTextStyle = "Off"
        occChkOrdered.Appearance = wdContentControlHidden
        occChkOrdered.LockContents = True
        
        occLblChkOrdered.DefaultTextStyle = "Off"
        
        
        occChkRestrictNOI.LockContents = False
        occChkRestrictNOI.DefaultTextStyle = "Heading 3"
        occChkRestrictNOI.Appearance = wdContentControlBoundingBox
        
        occLblChkRestrictNOI.DefaultTextStyle = "Heading 3"
        
        ToggelChkRestrictNOI
        
        
        occLblRscEligibleCustomers.DefaultTextStyle = "Heading 3"
        
        occTxtLblBtnAddCustomers.LockContents = False
        occTxtLblBtnAddCustomers.Range.Text = ""
        
        occTxtLblBtnClearCustomers.LockContents = False
        occTxtLblBtnClearCustomers.Range.Text = ""
        
        occRscEligibleCustomers.AllowInsertDeleteSection = True
    
        For Each occCustomer In ocsCustomers
            occCustomer.LockContents = False
            If occCustomer.Range.Text = "Disabled" Then: occCustomer.Range.Text = ""
            occCustomer.DefaultTextStyle = "ContentControlsSmall"
            occCustomer.Appearance = wdContentControlBoundingBox
        Next occCustomer
        
        
        occChkOnlyNewCustomers.LockContents = False
        occChkOnlyNewCustomers.DefaultTextStyle = "Heading 3"
        occChkOnlyNewCustomers.Appearance = wdContentControlBoundingBox
        
        occLblChkOnlyNewCustomers.DefaultTextStyle = "Heading 3"
        
        ToggelChkOnlyNewCustomers
    Else
        occChkInvoiced.LockContents = False
        occChkInvoiced.DefaultTextStyle = "Heading 3"
        occChkInvoiced.Appearance = wdContentControlBoundingBox
        
        occLblChkInvoiced.DefaultTextStyle = "Heading 3"
        
        occChkReceived.LockContents = False
        occChkReceived.DefaultTextStyle = "Heading 3"
        occChkReceived.Appearance = wdContentControlBoundingBox
        
        occLblChkReceived.DefaultTextStyle = "Heading 3"
        
        occChkOrdered.LockContents = False
        occChkOrdered.DefaultTextStyle = "Heading 3"
        occChkOrdered.Appearance = wdContentControlBoundingBox
        
        occLblChkOrdered.DefaultTextStyle = "Heading 3"
        
        
        occChkRestrictNOI.LockContents = False
        occChkRestrictNOI.Checked = False
        occChkRestrictNOI.DefaultTextStyle = "Off"
        occChkRestrictNOI.Appearance = wdContentControlHidden
        occChkRestrictNOI.LockContents = True
        
        occLblChkRestrictNOI.DefaultTextStyle = "Off"
        
        ToggelChkRestrictNOI
        
        If occChkPurchasing.Checked Or occChkOtherProgramType.Checked Then
            
            occLblRscEligibleCustomers.DefaultTextStyle = "Off"
            
            occTxtLblBtnAddCustomers.LockContents = False
            occTxtLblBtnAddCustomers.Range.Text = " "
            occTxtLblBtnAddCustomers.LockContents = True
            
            occTxtLblBtnClearCustomers.LockContents = False
            occTxtLblBtnClearCustomers.Range.Text = " "
            occTxtLblBtnClearCustomers.LockContents = True
            
            occRscEligibleCustomers.LockContentControl = False
            occRscEligibleCustomers.LockContents = False
            occRscEligibleCustomers.AllowInsertDeleteSection = True
            
            'Deleting all but two items in the RSC.
            'You can not delete the last item when there are only two, so we delete the first.
            Do While ocsRSCItems.Count >= 2
                For intIndex = 1 To ocsRSCItems(1).Range.ContentControls.Count
                    ocsRSCItems(1).Range.ContentControls(intIndex).LockContentControl = False
                Next intIndex
                ocsRSCItems.Item(1).Delete
            Loop

            occRscEligibleCustomers.LockContentControl = True
            occRscEligibleCustomers.LockContents = True
            occRscEligibleCustomers.AllowInsertDeleteSection = False

            For Each occCustomer In ocsCustomers
                occCustomer.LockContents = False
                occCustomer.Range.Text = "Disabled"
                occCustomer.DefaultTextStyle = "OffSmall"
                occCustomer.Appearance = wdContentControlHidden
                occCustomer.LockContents = True
                occCustomer.LockContentControl = True
            Next
            
            
            occChkOnlyNewCustomers.LockContents = False
            occChkOnlyNewCustomers.Checked = False
            occChkOnlyNewCustomers.DefaultTextStyle = "Off"
            occChkOnlyNewCustomers.Appearance = wdContentControlHidden
            
            occLblChkOnlyNewCustomers.DefaultTextStyle = "Off"
            
            ToggelChkOnlyNewCustomers
        Else
            occLblRscEligibleCustomers.DefaultTextStyle = "Heading 3"
            
            occTxtLblBtnAddCustomers.LockContents = False
            occTxtLblBtnAddCustomers.Range.Text = ""
            
            occTxtLblBtnClearCustomers.LockContents = False
            occTxtLblBtnClearCustomers.Range.Text = ""
            
            occRscEligibleCustomers.AllowInsertDeleteSection = True

            For Each occCustomer In ocsCustomers
                occCustomer.LockContents = False
                If occCustomer.Range.Text = "Disabled" Then: occCustomer.Range.Text = ""
                occCustomer.DefaultTextStyle = "ContentControlsSmall"
                occCustomer.Appearance = wdContentControlBoundingBox
            Next
            
            
            occChkOnlyNewCustomers.LockContents = False
            occChkOnlyNewCustomers.DefaultTextStyle = "Heading 3"
            occChkOnlyNewCustomers.Appearance = wdContentControlBoundingBox
            
            occLblChkOnlyNewCustomers.DefaultTextStyle = "Heading 3"
            
            ToggelChkOnlyNewCustomers
        End If 'occChkPurchasing.Checked Or occChkOtherProgramType.Checked
    End If 'occChkSales.Checked
    
    ToggelChkOtherProgramType
    ToggelAllowChkRestrictNOI
End Sub


' ------------------------------------------------------------------------------
' ToggelAllowChkRestrictNOI - Checks to see if NOI should be on or off.
'
' This check to see if neither purchasing or other program type fields are
' checked. If not it will enable the restrict NOI check box, if so it will
' un-check and disable it. Then it will set the NOI's subfields with its toggle.
' Finally we move the focus to the next field

Private Sub ToggelAllowChkRestrictNOI()
    Dim occChkSales As ContentControl
    Dim occChkPurchasing As ContentControl
    Dim occChkOtherProgramType As ContentControl
    Dim occChkRestrictNOI As ContentControl
    Dim occLblChkRestrictNOI As ContentControl
    
    Set occChkSales = ActiveDocument.SelectContentControlsByTag("chkSales")(1)
    Set occChkPurchasing = ActiveDocument.SelectContentControlsByTag("chkPurchasing")(1)
    Set occChkOtherProgramType = ActiveDocument.SelectContentControlsByTag("chkOtherProgramType")(1)
    Set occChkRestrictNOI = ActiveDocument.SelectContentControlsByTag("chkRestrictNOI")(1)
    Set occLblChkRestrictNOI = ActiveDocument.SelectContentControlsByTag("lblChkRestrictNOI")(1)
    
    If Not occChkPurchasing.Checked And Not occChkOtherProgramType.Checked Then
        occLblChkRestrictNOI.DefaultTextStyle = "Heading 3"
        
        occChkRestrictNOI.LockContents = False
        occChkRestrictNOI.DefaultTextStyle = "Heading 3"
    Else
        occLblChkRestrictNOI.DefaultTextStyle = "Off"
        
        occChkRestrictNOI.LockContents = False
        occChkRestrictNOI.Checked = False
        occChkRestrictNOI.DefaultTextStyle = "Off"
        occChkRestrictNOI.Appearance = wdContentControlHidden
        occChkRestrictNOI.LockContents = True
    End If
    
    ToggelChkRestrictNOI
    
    If occChkSales.Checked Or occChkPurchasing.Checked Then
        ActiveDocument.SelectContentControlsByTag("lstBuyer")(1).Range.Select
    Else
        ActiveDocument.SelectContentControlsByTag("txtOtherProgramType")(1).Range.Select
    End If
End Sub


' ------------------------------------------------------------------------------
' ToggelChkOtherProgramType - This toggles other program type fields on or off.
'
' This checks to make sure the sales and purchasing check boxes are empty. If so
' it will enable the other program type field. Or it will disable it if needed.

Private Sub ToggelChkOtherProgramType()
    Dim occChkSales As ContentControl
    Dim occChkPurchasing As ContentControl
    Dim occLblTxtOtherProgramType As ContentControl
    Dim occTxtOtherProgramType As ContentControl
    
    Set occChkSales = ActiveDocument.SelectContentControlsByTag("chkSales")(1)
    Set occChkPurchasing = ActiveDocument.SelectContentControlsByTag("chkPurchasing")(1)
    Set occLblTxtOtherProgramType = ActiveDocument.SelectContentControlsByTag("lblTxtOtherProgramType")(1)
    Set occTxtOtherProgramType = ActiveDocument.SelectContentControlsByTag("txtOtherProgramType")(1)

    If Not occChkSales.Checked And Not occChkPurchasing.Checked Then
        occLblTxtOtherProgramType.DefaultTextStyle = "Heading 4"
        
        occTxtOtherProgramType.LockContents = False
        If occTxtOtherProgramType.Range.Text = "Disabled" Then: occTxtOtherProgramType.Range.Text = ""
        occTxtOtherProgramType.DefaultTextStyle = "ContentControlsSmall"
        occTxtOtherProgramType.Appearance = wdContentControlBoundingBox
    Else
        occLblTxtOtherProgramType.DefaultTextStyle = "OffSmall"
        
        occTxtOtherProgramType.LockContents = False
        occTxtOtherProgramType.Range.Text = "Disabled"
        occTxtOtherProgramType.DefaultTextStyle = "OffSmall"
        occTxtOtherProgramType.Appearance = wdContentControlHidden
        occTxtOtherProgramType.LockContents = True
    End If
End Sub


' ------------------------------------------------------------------------------
' ToggelLstCategory - This toggles the if other program category on and off.
'
' This checks see if the program category is a placeholder or selected as other.
' If so we will turn on the field for typing a other category in, otherwise we
' will turn it off.

Private Sub ToggelLstCategory()
    Dim occLstCategory As ContentControl
    Dim occLblTxtOtherCategory As ContentControl
    Dim occTxtOtherCategory As ContentControl
    
    Set occLstCategory = ActiveDocument.SelectContentControlsByTag("lstCategory")(1)
    Set occLblTxtOtherCategory = ActiveDocument.SelectContentControlsByTag("lblTxtOtherCategory")(1)
    Set occTxtOtherCategory = ActiveDocument.SelectContentControlsByTag("txtOtherCategory")(1)

    If IsPlaceholder(occLstCategory) Or occLstCategory.Range.Text = "Other" Then
        occLblTxtOtherCategory.DefaultTextStyle = "Heading 4"
        
        occTxtOtherCategory.LockContents = False
        If occTxtOtherCategory.Range.Text = "Disabled" Then: occTxtOtherCategory.Range.Text = ""
        occTxtOtherCategory.DefaultTextStyle = "ContentControlsSmall"
        occTxtOtherCategory.Appearance = wdContentControlBoundingBox
    Else
        occLblTxtOtherCategory.DefaultTextStyle = "OffSmall"
        
        occTxtOtherCategory.LockContents = False
        occTxtOtherCategory.Range.Text = "Disabled"
        occTxtOtherCategory.DefaultTextStyle = "OffSmall"
        occTxtOtherCategory.Appearance = wdContentControlHidden
        occTxtOtherCategory.LockContents = True
    End If
End Sub


' ------------------------------------------------------------------------------
' ToggelChkEmailInvoice - This turns on or off our email fields.
'
' This checks see if the user checked the email box. If so it turns on all the
' subfields, if not it disables them. Then it selects the next field.

Private Sub ToggelChkEmailInvoice()
    Dim occChkEmailInvoice As ContentControl
    Dim occLblTxtToEmail1 As ContentControl
    Dim occTxtToEmail1 As ContentControl
    Dim occLblTxtToEmail2 As ContentControl
    Dim occTxtToEmail2 As ContentControl
    Dim occLblTxtToEmail3 As ContentControl
    Dim occTxtToEmail3 As ContentControl
    Dim occLblLstEmailFormat As ContentControl
    Dim occLstEmailFormat As ContentControl
    
    Set occChkEmailInvoice = ActiveDocument.SelectContentControlsByTag("chkEmailInvoice")(1)
    Set occLblTxtToEmail1 = ActiveDocument.SelectContentControlsByTag("lblTxtToEmail1")(1)
    Set occTxtToEmail1 = ActiveDocument.SelectContentControlsByTag("txtToEmail1")(1)
    Set occLblTxtToEmail2 = ActiveDocument.SelectContentControlsByTag("lblTxtToEmail2")(1)
    Set occTxtToEmail2 = ActiveDocument.SelectContentControlsByTag("txtToEmail2")(1)
    Set occLblTxtToEmail3 = ActiveDocument.SelectContentControlsByTag("lblTxtToEmail3")(1)
    Set occTxtToEmail3 = ActiveDocument.SelectContentControlsByTag("txtToEmail3")(1)
    Set occLblLstEmailFormat = ActiveDocument.SelectContentControlsByTag("lblLstEmailFormat")(1)
    Set occLstEmailFormat = ActiveDocument.SelectContentControlsByTag("lstEmailFormat")(1)

    If occChkEmailInvoice.Checked Then
        occLblTxtToEmail1.DefaultTextStyle = "Heading 4"
        
        occTxtToEmail1.LockContents = False
        occTxtToEmail1.Range.Text = ""
        occTxtToEmail1.DefaultTextStyle = "ContentControlsSmall"
        occTxtToEmail1.Appearance = wdContentControlBoundingBox
        
        occLblTxtToEmail2.DefaultTextStyle = "Heading 4"
        
        occTxtToEmail2.LockContents = False
        occTxtToEmail2.Range.Text = ""
        occTxtToEmail2.DefaultTextStyle = "ContentControlsSmall"
        occTxtToEmail2.Appearance = wdContentControlBoundingBox
        
        occLblTxtToEmail3.DefaultTextStyle = "Heading 4"
        
        occTxtToEmail3.LockContents = False
        occTxtToEmail3.Range.Text = ""
        occTxtToEmail3.DefaultTextStyle = "ContentControlsSmall"
        occTxtToEmail3.Appearance = wdContentControlBoundingBox
        
        occLblLstEmailFormat.DefaultTextStyle = "Heading 4"
        
        occLstEmailFormat.LockContents = False
        Call ResetList(occLstEmailFormat.Tag, , True)
        occLstEmailFormat.DefaultTextStyle = "ContentControlsSmall"
        occLstEmailFormat.Appearance = wdContentControlBoundingBox
    Else
        occLblTxtToEmail1.DefaultTextStyle = "OffSmall"
        
        occTxtToEmail1.LockContents = False
        occTxtToEmail1.Range.Text = "Disabled"
        occTxtToEmail1.DefaultTextStyle = "OffSmall"
        occTxtToEmail1.Appearance = wdContentControlHidden
        occTxtToEmail1.LockContents = True
        
        occLblTxtToEmail2.DefaultTextStyle = "OffSmall"
        
        occTxtToEmail2.LockContents = False
        occTxtToEmail2.Range.Text = "Disabled"
        occTxtToEmail2.DefaultTextStyle = "OffSmall"
        occTxtToEmail2.Appearance = wdContentControlHidden
        occTxtToEmail2.LockContents = True
        
        occLblTxtToEmail3.DefaultTextStyle = "OffSmall"
        
        occTxtToEmail3.LockContents = False
        occTxtToEmail3.Range.Text = "Disabled"
        occTxtToEmail3.DefaultTextStyle = "OffSmall"
        occTxtToEmail3.Appearance = wdContentControlHidden
        occTxtToEmail3.LockContents = True
        
        occLblLstEmailFormat.DefaultTextStyle = "OffSmall"
        
        occLstEmailFormat.LockContents = False
        Call ResetList(occLstEmailFormat.Tag, "Disabled", True)
        occLstEmailFormat.DefaultTextStyle = "OffSmall"
        occLstEmailFormat.Appearance = wdContentControlHidden
        occLstEmailFormat.LockContents = True
    End If
  
    If occChkEmailInvoice.Checked Then
        occTxtToEmail1.Range.Select
    Else
        If ActiveDocument.SelectContentControlsByTag("chkPayBuyingGroup")(1).Checked Then
            ActiveDocument.SelectContentControlsByTag("txtBuyingGroup")(1).Range.Select
        Else
            ActiveDocument.SelectContentControlsByTag("lstGLAccount")(1).Range.Select
        End If
    End If
End Sub


' ------------------------------------------------------------------------------
' ToggelChkPayBuyingGroup - This turns on or off our buying group fields.
'
' This checks see if the user checked the buying box. If so it turns on all the
' subfields, if not it disables them.

Private Sub ToggelChkPayBuyingGroup()
    Dim occChkPayBuyingGroup As ContentControl
    Dim occLblTxtBuyingGroup As ContentControl
    Dim occTxtBuyingGroup As ContentControl
    
    Set occChkPayBuyingGroup = ActiveDocument.SelectContentControlsByTag("chkPayBuyingGroup")(1)
    Set occLblTxtBuyingGroup = ActiveDocument.SelectContentControlsByTag("lblTxtBuyingGroup")(1)
    Set occTxtBuyingGroup = ActiveDocument.SelectContentControlsByTag("txtBuyingGroup")(1)

    If occChkPayBuyingGroup.Checked Then
        occLblTxtBuyingGroup.DefaultTextStyle = "Heading 4"
    
        occTxtBuyingGroup.LockContents = False
        If occTxtBuyingGroup.Range.Text = "Disabled" Then: occTxtBuyingGroup.Range.Text = ""
        occTxtBuyingGroup.DefaultTextStyle = "ContentControlsSmall"
        occTxtBuyingGroup.Appearance = wdContentControlBoundingBox
        occTxtBuyingGroup.Range.Select
    Else
        occLblTxtBuyingGroup.DefaultTextStyle = "OffSmall"
    
        occTxtBuyingGroup.LockContents = False
        occTxtBuyingGroup.Range.Text = "Disabled"
        occTxtBuyingGroup.DefaultTextStyle = "OffSmall"
        occTxtBuyingGroup.Appearance = wdContentControlHidden
        occTxtBuyingGroup.LockContents = True
        ActiveDocument.SelectContentControlsByTag("lstGLAccount")(1).Range.Select
    End If
End Sub


' ------------------------------------------------------------------------------
' ToggelLstRebateType - This turns on or off our buying group fields.
'
' This checks see if the user selected regular rebate or left it as the
' placeholder. If so it will turn on the paid by field, if not it
' disables that field. It also checks all the RSC items payment types
' and sets them to the proper states as well.

Private Sub ToggelLstRebateType()
    Dim occLstRebateType As ContentControl
    Dim occLblLstPaidBy As ContentControl
    Dim occLstPaidBy As ContentControl
    Dim ocsItems As ContentControls
    Dim occItem As ContentControl
    
    Set occLstRebateType = ActiveDocument.SelectContentControlsByTag("lstRebateType")(1)
    Set occLblLstPaidBy = ActiveDocument.SelectContentControlsByTag("lblLstPaidBy")(1)
    Set occLstPaidBy = ActiveDocument.SelectContentControlsByTag("lstPaidBy")(1)
    Set ocsItems = ActiveDocument.SelectContentControlsByTag("lstItemPaidBy")

    If IsPlaceholder(occLstRebateType) Or occLstRebateType.Range.Text = "Regular" Then
        occLblLstPaidBy.DefaultTextStyle = "Heading 4"
        
        occLstPaidBy.LockContents = False
        If occLstPaidBy.Range.Text = "Disabled" Then: Call ResetList(occLstPaidBy.Tag, , True)
        occLstPaidBy.DefaultTextStyle = "ContentControlsSmall"
        occLstPaidBy.Appearance = wdContentControlBoundingBox
        
        For Each occItem In ocsItems
            If occItem.Range.Text = "Disabled" Then
                occItem.LockContents = False
                occItem.Type = wdContentControlText
                occItem.Range.Text = ""
                occItem.DefaultTextStyle = "ContentControlsList"
                occItem.Appearance = wdContentControlBoundingBox
                occItem.Type = wdContentControlDropdownList
            End If
        Next occItem
    Else
        occLblLstPaidBy.DefaultTextStyle = "OffSmall"
        
        occLstPaidBy.LockContents = False
        Call ResetList(occLstPaidBy.Tag, "Disabled", True)
        occLstPaidBy.DefaultTextStyle = "OffSmall"
        occLstPaidBy.Appearance = wdContentControlHidden
        occLstPaidBy.LockContents = True
        
        For Each occItem In ocsItems
            occItem.LockContents = False
            occItem.Type = wdContentControlText
            occItem.Range.Text = "Disabled"
            occItem.DefaultTextStyle = "OffList"
            occItem.Appearance = wdContentControlHidden
            occItem.Type = wdContentControlDropdownList
            occItem.LockContents = True
        Next occItem
    End If
End Sub


' ------------------------------------------------------------------------------
' ToggelChkRestrictNOI - This turns on or off the Restrict NOI fields.
'
' This checks see if the user checked the restrict NOI box. If so it
' enables the NOI amount and unit fields, while clearing them if
' they were disabled. If not it disables these fields.

Private Sub ToggelChkRestrictNOI()
    Dim occChkRestrictNOI As ContentControl
    Dim occLblTxtNOIAmount As ContentControl
    Dim occTxtNOIAmount As ContentControl
    Dim occLblLstNOIUnit As ContentControl
    Dim occLstNOIUnit As ContentControl
    
    Set occChkRestrictNOI = ActiveDocument.SelectContentControlsByTag("chkRestrictNOI")(1)
    Set occLblTxtNOIAmount = ActiveDocument.SelectContentControlsByTag("lblTxtNOIAmount")(1)
    Set occTxtNOIAmount = ActiveDocument.SelectContentControlsByTag("txtNOIAmount")(1)
    Set occLblLstNOIUnit = ActiveDocument.SelectContentControlsByTag("lblLstNOIUnit")(1)
    Set occLstNOIUnit = ActiveDocument.SelectContentControlsByTag("lstNOIUnit")(1)

    If occChkRestrictNOI.Checked Then
        occLblTxtNOIAmount.DefaultTextStyle = "Heading 4"
        
        occTxtNOIAmount.LockContents = False
        If occTxtNOIAmount.Range.Text = "Disabled" Then: occTxtNOIAmount.Range.Text = ""
        occTxtNOIAmount.DefaultTextStyle = "ContentControlsSmall"
        occTxtNOIAmount.Appearance = wdContentControlBoundingBox
        
        occLblLstNOIUnit.DefaultTextStyle = "Heading 4"
        
        occLstNOIUnit.LockContents = False
        If occLstNOIUnit.Range.Text = "Disabled" Then: Call ResetList(occLstNOIUnit.Tag, , True)
        occLstNOIUnit.DefaultTextStyle = "ContentControlsSmall"
        occLstNOIUnit.Appearance = wdContentControlBoundingBox
    Else
        occLblTxtNOIAmount.DefaultTextStyle = "OffSmall"
        
        occTxtNOIAmount.LockContents = False
        occTxtNOIAmount.Range.Text = "Disabled"
        occTxtNOIAmount.DefaultTextStyle = "OffSmall"
        occTxtNOIAmount.Appearance = wdContentControlHidden
        occTxtNOIAmount.LockContents = True
        
        occLblLstNOIUnit.DefaultTextStyle = "OffSmall"
        
        occLstNOIUnit.LockContents = False
        Call ResetList(occLstNOIUnit.Tag, "Disabled", True)
        occLstNOIUnit.DefaultTextStyle = "OffSmall"
        occLstNOIUnit.Appearance = wdContentControlHidden
        occLstNOIUnit.LockContents = True
    End If
End Sub


' ------------------------------------------------------------------------------
' ToggelChkOnlyNewCustomers - This turns on or off new customers only fields.
'
' This checks see if the user checked the only new customers box. If so it
' turns on the date field. If not it disables it.

Private Sub ToggelChkOnlyNewCustomers()
    Dim occChkOnlyNewCustomers As ContentControl
    Dim occDtpNewSince As ContentControl
    Dim occLblDtpNewSince As ContentControl
    
    Set occChkOnlyNewCustomers = ActiveDocument.SelectContentControlsByTag("chkOnlyNewCustomers")(1)
    Set occDtpNewSince = ActiveDocument.SelectContentControlsByTag("dtpNewSince")(1)
    Set occLblDtpNewSince = ActiveDocument.SelectContentControlsByTag("lblDtpNewSince")(1)

    If occChkOnlyNewCustomers.Checked Then
        occLblDtpNewSince.DefaultTextStyle = "Heading 4"
        
        occDtpNewSince.LockContents = False
        occDtpNewSince.Range.Text = ""
        occDtpNewSince.DefaultTextStyle = "ContentControlsSmall"
        occDtpNewSince.Appearance = wdContentControlBoundingBox
    Else
        occLblDtpNewSince.DefaultTextStyle = "OffSmall"
        
        occDtpNewSince.LockContents = False
        occDtpNewSince.Range.Text = "Disabled"
        occDtpNewSince.DefaultTextStyle = "OffSmall"
        occDtpNewSince.Appearance = wdContentControlHidden
        occDtpNewSince.LockContents = True
    End If
    
    If occChkOnlyNewCustomers.Checked Then
        occDtpNewSince.Range.Select
    Else
        ActiveDocument.SelectContentControlsByTag("txtNotes")(1).Range.Select
    End If
End Sub
