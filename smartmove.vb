Public Sub MoveConversation2()
    ' Error handling setup
    On Error GoTo ErrorHandler
    
    ' Declaration of variables
    Dim objExplorer As Outlook.Explorer
    Dim objSelection As Outlook.Selection
    Dim objMail As Outlook.MailItem
    Dim objConversation As Outlook.Conversation
    Dim objTable As Outlook.Table
    Dim objRow As Outlook.Row
    Dim destinationFolder As Outlook.Folder
    Dim earliestMail As Outlook.MailItem
    
    ' Start of the process
    Debug.Print "Step 1: Object initialization complete."
    
    ' Get the current Explorer window and the selected items
    Set objExplorer = Application.ActiveExplorer
    Set objSelection = objExplorer.Selection
    
    Debug.Print "Step 2: Selection objects initialized."
    
    ' Check if there is at least one item selected
    If objSelection.Count > 0 Then
        Debug.Print "Step 3: Selection is more than 0."
        
        ' Ensure the selected item is a MailItem
        If TypeOf objSelection.Item(1) Is MailItem Then
            Debug.Print "Step 4: Selection is MailItem."
            
            ' Get the first selected mail item and its conversation
            Set objMail = objSelection.Item(1)
            Set objConversation = objMail.GetConversation

            ' Check if the conversation is not null
            If Not objConversation Is Nothing Then
                Debug.Print "Step 5: objConversation is not nothing."
                
                ' Retrieve the earliest mail in the conversation
                Set objTable = objConversation.GetTable
                objTable.Sort "[ReceivedTime]", False
                Set objRow = objTable.GetNextRow

                ' Loop through the conversation to find the earliest mail
                Do Until objRow Is Nothing
                    Set earliestMail = Application.Session.GetItemFromID(objRow("EntryID"))
                    If TypeOf earliestMail Is MailItem Then
                        Exit Do
                    End If
                    Set objRow = objTable.GetNextRow
                Loop

                ' Check if the earliest mail is not null
                If Not earliestMail Is Nothing Then
                    ' Set the folder where the earliest mail is located
                    Set destinationFolder = earliestMail.Parent
                    Debug.Print "Destination Folder: " & destinationFolder.Name
                    
                    ' Move the selected email to the folder of the earliest email
                    objMail.Move destinationFolder

                    
                End If
            End If
        End If
    End If
    Exit Sub

ErrorHandler:
    ' Display any errors that occur
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume Next
End Sub
