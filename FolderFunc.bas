''' デフォルトでメールで使用するフォルダを取得します。
''' return: Folderコレクション
Public Function GetDefaultFolderByMail() As Collection
    Dim folderList As New Collection
    Dim ns As NameSpace
    Set ns = Application.GetNamespace("MAPI")
    
    Call folderList.Add(ns.GetDefaultFolder(olFolderInbox)) '受信トレイ
    Call folderList.Add(ns.GetDefaultFolder(olFolderSentMail)) '送信済み
    Call folderList.Add(ns.GetDefaultFolder(olFolderDeletedItems)) '削除済み
    
    Set GetDefaultFolderByMail = folderList
End Function

''' デフォルトでメールで使用するフォルダを取得します。
''' folderPath: 取得するフォルダパス　例："\\Mailbox - mailAddress\Inbox\Customers"
''' return: Folderコレクション
Public Function GetFolder(ByVal folderPath As String) As Outlook.folder
    Dim TestFolder As Outlook.folder
    Dim FoldersArray As Variant
    Dim i As Integer
 
    On Error GoTo GetFolder_Error
    If Left(folderPath, 2) = "\\" Then
        folderPath = Right(folderPath, Len(folderPath) - 2)
    End If
    
    'Convert folderpath to array
    FoldersArray = Split(folderPath, "\")
    Set TestFolder = Application.Session.folders.item(FoldersArray(0))
    If Not TestFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.folders
            Set SubFolders = TestFolder.folders
            Set TestFolder = SubFolders.item(FoldersArray(i))
            If TestFolder Is Nothing Then
                Set GetFolder = Nothing
            End If
        Next
    End If
     
    Set GetFolder = TestFolder
    Exit Function
 
GetFolder_Error:
    Set GetFolder = Nothing
    Exit Function
End Function

Public Function GetCurrentItem() As Object
    Set GetCurrentItem = Nothing

    Dim olApp As Outlook.Application
    Set olApp = Application
    
    On Error Resume Next

    Select Case TypeName(olApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = olApp.ActiveExplorer.Selection.item(1)
        Case "Inspector"
            Set GetCurrentItem = olApp.ActiveInspector.currentItem
    End Select
    

    Set olApp = Nothing
End Function

Public Function GetDictionaryInstance() As Object
    Set GetDictionaryInstance = CreateObject("Scripting.Dictionary")
End Function
