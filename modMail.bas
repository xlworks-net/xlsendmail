'//
'//
'// #     # #     #  #  #  ####  ####  #   #  ####      #     # #### #####
'//  #   #  #     #  #  # #    # #   # #  #  #          ##    # #      #
'//   # #   #     #  #  # #    # #   # # #   #          # #   # #      #
'//    #    #     #  #  # #    # ####  ##     ###       #  #  # ###    #
'//   # #   #     #  #  # #    # #  #  # #       #      #   # # #      #
'//  #   #  #     #  #  # #    # #   # #  #      #      #    ## #      #
'// #     # #####  ## ##   ####  #   # #   # ####   #   #     # ####   #
'//
'//
'// Send mail using Outlook
'// 2020.11.22 xlworks.net
'//
'// Original source: http: //www.rondebruin.nl/win/section1.htm
'//
Option Explicit

Sub checkSenderList(ByVal argSendType As String)
    '//If there are more than one email account, popup is displayed so that users can select it.
    
Dim i As Long
Dim arrData() As String

   On Error GoTo ErrHandler
    
    '//If an error occurs when creating an Outlook object, go ahead and check if OutApp is Nothing, and if it is Nothing, treat it as an error.
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    
    If OutApp Is Nothing Then
        MsgBox MSG_OUTLOOK_CANNOT_BE_USED
        Exit Sub
    End If
    On Error GoTo 0

    
    If OutApp.Session.Accounts.Count > 1 Then
        '//Since there are more than one email account, popup is displayed.
        
        ReDim arrData(0 To OutApp.Session.Accounts.Count - 1, 0 To 1)
            
        For i = 0 To OutApp.Session.Accounts.Count - 1
            arrData(i, 0) = i + 1
            arrData(i, 1) = OutApp.Session.Accounts.item(i + 1)
        Next i
        
        With frmChooseSender.cbxSender
            .List = arrData
            .ListIndex = 0
        End With
        
        frmChooseSender.txtSendType = argSendType
        frmChooseSender.Show
               
    Else
        '//If there is only one email account, send mail without popup
        sendEmail argSendType:=argSendType, argSenderIndex:=1
    End If
    
   
NormalEnd:

    Set OutApp = Nothing
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
Exit Sub

ErrHandler:

    If Err.Number = "287" Then
        MsgBox MSG_ERROR_OCCURRED_CHECK_ANTIVIRUS_OPTION & vbNewLine & vbNewLine & "Error code : " & Err.Number & vbNewLine & Err.Description
    ElseIf Err.Number <> 0 Then
        MsgBox MSG_ERROR_OCCURRED_CHECK_BELOW & vbNewLine & vbNewLine & "Error code : " & Err.Number & vbNewLine & Err.Description
    End If
    
    Resume NormalEnd
  
End Sub
Sub sendEmail(ByVal argSendType As String, ByVal argSenderIndex As Integer)

'// * Note: Error when sending Outlook mail using VBA.
'//
'// If you do not have an antivirus program installed on your pc
'// (even if it is installed but it is not running), you may see an Outlook Security Warning window.
'// If you select Deny in the security alert window, you will get a 287 error.
'// If Outlook is not running, the Outlook Security Warning window does not appear
'// and immediately a 287 error occurs (depending on the O / S version, it may be slightly different)
'//
'// * Outlook Security related reference sites:
'//   http://www.outlookcode.com/article.aspx?ID=52
'//   http://www.rondebruin.nl/win/s1/security.htm
'//   https://msdn.microsoft.com/en-us/library/ms778202.aspx
Dim OutInspector As Object
Dim OutMail As Object
Dim oOutlook As Object

Dim sh As Worksheet
Dim cell As Range, FileCell As Range, rng As Range
Dim rngMailList As Range
Dim strSubject As String
Dim strMsgBody As String
Dim strSignature As String
Dim strFontSelect1 As String
Dim strFontSelect2 As String
Dim sendCount As Long
Dim bOriginatorDeliveryReportRequested  As Boolean
Dim bReadReceiptRequested As Boolean
Dim bSignatureUse As Boolean
Dim userDeferredDeliveryTime As Date
Dim userDeferredDeliveryTimeString As Variant

    
On Error GoTo ErrHandler

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

  
    strSubject = Range("RNG_SUBJECT").Value
    strMsgBody = Range("RNG_BODY").Value
    
    '//Request a delivery receipt
    If Range("RNG_ORIGINATOR_DELIVERY_REPORT_REQUESTED").Value = "YES" Then
        bOriginatorDeliveryReportRequested = True
    End If
    
    '//Request a read receipt
    If Range("RNG_READ_RECEIPT_REQUESTED").Value = "YES" Then
        bReadReceiptRequested = True
    End If
    
    '//Signature
    If Range("RNG_SIGNATURE_USE").Value = "YES" Then
        bSignatureUse = True
        
        '//Added @2020.11.21
        '//test if Outlook is running.
        '//If Outlook is not running, you cannot use the signature feature.
        On Error Resume Next
        Set oOutlook = GetObject(, "Outlook.Application")
        On Error GoTo 0

        If oOutlook Is Nothing Then
            MsgBox MSG_RUN_OUTLOOK_FIRST
            Exit Sub
        End If
        
    End If
    
    '//font type of mail body
    If Range("RNG_FONT_SELECT").Value = "YES" Then
        strFontSelect1 = "<p style='font-family:" & Sheets(MAIL_CONTENTS_SHEET).cbxFontLists.Value & ";font-size:" & Sheets(MAIL_CONTENTS_SHEET).cbxFontSize.Value & "pt'>"
        strFontSelect2 = "</p>"
    End If


    Set sh = Sheets(MAIL_LIST_SHEET)
    
    If argSendType = "SEND" Then
        If MsgBox(MSG_WANT_TO_MAIL, vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    

    '// About Excel Range.SpecialCells Method
    '// https://docs.microsoft.com/en-us/office/vba/api/excel.range.specialcells
    '// xlCellTypeVisible - All visible cells
    '//For Each cell In sh.Columns(TO_ADDRESS_COL).Cells.SpecialCells(xlCellTypeConstants)
    '//
    '//In sendEmail Procedure, when using xlCellTypeConstants to retrieve e-mail address using vlookup function,
    '//the result of function is not recognized. So I changed xlCellTypeConstants (the cell containing constants) to xlCellTypeVisible (all visible cells).
    For Each cell In sh.Columns(TO_ADDRESS_COL).Cells.SpecialCells(xlCellTypeVisible)
        If VBA.UCase(cell.Offset(0, -1)) = "YES" Then

            'Attach file range
            Set rng = sh.Cells(cell.Row, 1).Range(ATTACH_FOLDER_START & "1:" & ATTACH_FOLDER_END & "1")
    
            If cell.Value Like "?*@?*.?*" Then
                Set OutMail = OutApp.CreateItem(0)
    
                With OutMail
                    .ReadReceiptRequested = bReadReceiptRequested
                    .OriginatorDeliveryReportRequested = bOriginatorDeliveryReportRequested
                    .To = cell.Value
                    .CC = cell.Offset(, CC_ADDRESS_COL - TO_ADDRESS_COL).Value
                    .BCC = cell.Offset(, BCC_ADDRESS_COL - TO_ADDRESS_COL).Value
                    
                    '//2019.4.18 Changed to allow deferred delivery time to be specified as an substitute value for each mail
                    '//Deferred delivery time
                    userDeferredDeliveryTimeString = Range("RNG_DEFERRED_DELIVERY_TIME").Value
                    userDeferredDeliveryTimeString = replaceContents(userDeferredDeliveryTimeString, sh.Cells(cell.Row, 1).Range(SUBSTITUTE_VALUE_START & "1:" & SUBSTITUTE_VALUE_END & "1"))
                    
                    If userDeferredDeliveryTimeString > "" Then
                        If IsDate(userDeferredDeliveryTimeString) Then
                            If userDeferredDeliveryTimeString >= Now() Then
                                userDeferredDeliveryTime = userDeferredDeliveryTimeString
                            Else
                                MsgBox MSG_DEFERRED_DELIVERY_TIME_ENTERED_AFTER_CURRENT_TIME
                                Exit Sub
                            End If
                        Else
                            MsgBox MSG_DEFERRED_DELIVERY_TIME_FORMAT_ERROR
                            Exit Sub
                        End If
                    End If
                    
                    
                    
                    '//2018.6.23 Fixed a bug when a message exists in sent box but the message can not be sent
                    If userDeferredDeliveryTime > 0 Then  '//when not empty "deferred time delivery" field
                        .deferredDeliveryTime = userDeferredDeliveryTime
                    End If
                    
                    '//Added the feature to select accounts when using multiple accounts in Outlook
                    Set .SendUsingAccount = OutApp.Session.Accounts.item(argSenderIndex)
                    
                    '//Change the subject to the substitute value that the user entered in Excel
'                    .Subject = replaceContents(strSubject, Range(cell.Offset(, 0), cell.Offset(, 24)))
                    .Subject = replaceContents(strSubject, sh.Cells(cell.Row, 1).Range(SUBSTITUTE_VALUE_START & "1:" & SUBSTITUTE_VALUE_END & "1"))

                    '//Call Getinspector to make the signature available only if you choose to use it.
                    If bSignatureUse Then
                        Set OutInspector = OutMail.Getinspector
                    End If
                                
                    '//Change the mail body to the substitute value that the user entered in Excel
'                    .htmlbody = "<html><body>" & strFontSelect1 & replaceContents(strMsgBody, Range(cell.Offset(, 0), cell.Offset(, 24))) & .htmlbody & strFontSelect2 & "</body></html>"
                    .htmlbody = "<html><body>" & strFontSelect1 & replaceContents(strMsgBody, sh.Cells(cell.Row, 1).Range(SUBSTITUTE_VALUE_START & "1:" & SUBSTITUTE_VALUE_END & "1")) & .htmlbody & strFontSelect2 & "</body></html>"
                                
                    '//If there is a value in the path of the attachment, process the attachment.
                    
                    '//2019.4.19 "xlCellTypeConstants" changed to "xlCellTypeVisible" because there was a problem that the path was not recognized by vba when it was entered as a formula.
                    If Application.WorksheetFunction.CountA(rng) > 0 Then
                        '//For Each FileCell In rng.SpecialCells(xlCellTypeConstants)
                        For Each FileCell In rng.SpecialCells(xlCellTypeVisible)
                            If VBA.Trim(FileCell) <> "" Then
                                If Dir(FileCell.Value) <> "" Then
                                    '//.Attachments.Add FileCell.Value
                                    .Attachments.Add FileCell.Text
                                End If
                            End If
                        Next FileCell
                    End If

                    '//.Display  'Or use Send
                    If argSendType = "SEND" Then
                        .send
                    Else
                        .display
                        Set OutMail = Nothing
                        GoTo NormalEnd
                    End If
                    
                End With
                
                '//Increase count when mail is sent.
                sendCount = sendCount + 1
            
                Set OutMail = Nothing
            
            End If
        End If
    Next cell
    

    
    If sendCount = 0 Then
        MsgBox MSG_MAIL_NOT_SENT_BECAUSE_NO_TARGET
    Else
        MsgBox MSG_MAIL_SENT & vbNewLine & _
               MSG_MAIL_SENT_PART1 & sh.Columns("A").Cells.SpecialCells(xlCellTypeConstants).Count - 1 & MSG_MAIL_SENT_PART2 & sendCount & MSG_MAIL_SENT_PART3
    End If
    
NormalEnd:

    Set OutApp = Nothing
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
Exit Sub


ErrHandler:

    If Err.Number = "287" Then
        MsgBox MSG_ERROR_OCCURRED_CHECK_ANTIVIRUS_OPTION & vbNewLine & vbNewLine & "Error code : " & Err.Number & vbNewLine & Err.Description
    ElseIf Err.Number <> 0 Then
        MsgBox MSG_ERROR_OCCURRED_CHECK_BELOW & vbNewLine & vbNewLine & "Error code : " & Err.Number & vbNewLine & Err.Description
    End If
    
    Resume NormalEnd
  
End Sub