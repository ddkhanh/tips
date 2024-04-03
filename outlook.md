### 1 Setup Outlook to allow you to create a new email which based on the selected email
#### 1.1 Scripts
```vba
Sub CreateEmailFromSelected()
    Dim objSelection As Outlook.Selection
    Dim objItem As Object
    Dim objMail As Outlook.MailItem
    Dim objAttachment As Outlook.Attachment
    Dim objNamespace As Outlook.NameSpace
    Dim objRecipient As Outlook.recipient
    
    ' Get the current selection
    Set objSelection = Outlook.Application.ActiveExplorer.Selection
    
    ' Check if something is selected
    If Not objSelection Is Nothing Then
        ' Loop through each selected item
        For Each objItem In objSelection
            ' Check if the selected item is a mail item
            If TypeOf objItem Is Outlook.MailItem Then
                ' Create a new email
                Set objMail = Outlook.Application.CreateItem(olMailItem)
                With objMail
                    .subject = objItem.subject
                    .HTMLBody = objItem.HTMLBody

                    ' Set the recipients (To)
                    For Each objRecipient In objItem.Recipients
                        If objRecipient.Type = olTo Then
                            objMail.Recipients.Add objRecipient.Address
                        End If
                    Next objRecipient
                
                    ' Set the recipients (CC)
                    For Each objRecipient In objItem.Recipients
                        If objRecipient.Type = olCC Then
                            objMail.CC = objMail.CC & ";" & objRecipient.Address
                        End If
                    Next objRecipient
                
                    objMail.Recipients.ResolveAll ' Resolve recipients
                    
                    ' Copy attachments, if any
                    For Each objAttachment In objItem.Attachments
                        objAttachment.SaveAsFile Environ("Temp") & "\" & objAttachment.FileName
                        objMail.Attachments.Add Environ("Temp") & "\" & objAttachment.FileName
                    Next objAttachment
                End With

                ' Display the new email
                objMail.Display
            End If
        Next objItem
    Else
        MsgBox "No emails selected.", vbExclamation
    End If
End Sub
```
#### 1.2 Demo

![image](https://github.com/ddkhanh/tips/assets/5151868/c3dce8df-fbc9-433a-9a13-eec053368624)
