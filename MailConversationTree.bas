
'検索対象のフォルダ 文字列型でカンマ区切りにする
'受信トレイ、送信済み、削除済みフォルダはデフォルトで検索されるので追加しないでください
Private Const SEARCH_FOLDER_PATH_LIST = "YourFolderPathList"

Private MailTreeNode As treeNode

''' Main関数
Public Sub FeatureMailConversationTree()
    Dim currentItem As mailItem
    Set currentItem = mail.GetCurrentMailItem()
    
    If currentItem Is Nothing Then
        Call MsgBox("メールを選択してください")
    End If
    ' 会話を取得
    Dim conversationItem As conversation
    Set conversationItem = currentItem.GetConversation()
    Dim convesationRootItems As SimpleItems
    Set convesationRootItems = conversationItem.GetRootItems()
    
    Set MailTreeNode = New treeNode
    Set MailTreeNode.Data = convesationRootItems.item(1)
    
    ' MailTreeNodeを構築するためのサブプロシージャ
    Call ProcessConversationRootItems(GetMailItemOnly(convesationRootItems), MailTreeNode)
    
    Call ConversationTreeForm.Show(False) 'モードレスじゃないとメールが開けない
    Call ConversationTreeForm.DrawTree(MailTreeNode)
    
    
End Sub


Private Sub ProcessConversationRootItems(ByVal rootMailItems As Collection, ByRef parentNode As treeNode)
    Dim argsItem As mailItem
    Dim argsTreeNode As treeNode
    
    ' 既定の初期値を設定
    Set argsItem = rootMailItems.item(1)
    Set argsTreeNode = parentNode
    
    Select Case rootMailItems.count
        Case 1
            ' アイテムが1つの場合、何も変更せずにそのまま進む
        Case 2
            ' アイテムが2つの場合、ConversationIndexを比較して処理を進める
            If rootMailItems.item(1).ConversationIndex = rootMailItems.item(2).ConversationIndex Then
                Dim nextNode As treeNode
                Set nextNode = New treeNode
                Call parentNode.AddChild(nextNode)
                Set nextNode.Data = rootMailItems.item(2)
                
                Set argsItem = rootMailItems.item(2)
                Set argsTreeNode = nextNode
            End If
            
        Case Else
            ' 想定外のケースの場合、デバッグ出力して終了
            Call Err.Raise(1111, , "Mailアイテムの数が異常です。処理を終了します")
            Exit Sub
    End Select
    
    Call BuildConversationTree(argsItem, argsTreeNode, argsItem.ConversationID)
End Sub



Private Function BuildConversationTree(ByVal rootMailItem As Outlook.mailItem, ByRef parentTreeNode As treeNode, ByVal originalConversationId As String) As treeNode
    Dim conversation As Outlook.conversation
    
    ' 会話を取得して子アイテムを処理
    Set conversation = rootMailItem.GetConversation
    If conversation Is Nothing Then
        Exit Function
    End If
    
    Dim childItems As SimpleItems
    Set childItems = conversation.GetChildren(rootMailItem)

    ' ConversationIndexでグルーピングする
    Dim conversationIndexGroups As Object  'ディクショナリ
    Set conversationIndexGroups = mail.GroupMailItemsByConversationIndex(mail.GetMailItemOnly(childItems))
    
    Call ProcessConversationIndexGroups(conversationIndexGroups, parentTreeNode, originalConversationId)
    
    ' ConversationID(件名)が変わった返信がないか調べ、あれば再起処理
    Dim replyMailItems As Collection
    Set replyMailItems = GetReplyMailItemsInFolders(GetInternetMessageId(rootMailItem), GetSearchFolders())
    
    ' 今と同じ会話Idのメールアイテムを削除し、新しいコレクションを取得
    Set replyMailItems = FilterMailItemsByDifferentConversationId(replyMailItems, originalConversationId)
    
    ' ConversationIDでグルーピングされたメールアイテムを処理
    Dim conversationIdGroups As Object  'ディクショナリ
    Set conversationIdGroups = GroupMailItemsByConversationId(replyMailItems)
    
    Dim conversationIdKey As Variant
    For Each conversationIdKey In conversationIdGroups.keys
        Dim groupedMailItems As Collection
        Set groupedMailItems = conversationIdGroups(conversationIdKey)
        
        Dim conversationIndexGroupsOther As Object  'ディクショナリ
        Set conversationIndexGroupsOther = GroupMailItemsByConversationIndex(groupedMailItems)
        
        Call ProcessConversationIndexGroups(conversationIndexGroupsOther, parentTreeNode, originalConversationId)
    Next
    
    Set BuildConversationTree = parentTreeNode
End Function

''' conversationIndexでグループ化されたディクショナリに対して処理
Private Sub ProcessConversationIndexGroups(ByVal conversationIndexGroups As Object, ByRef parentTreeNode As treeNode, ByVal originalConversationId As String)
    Dim conversationIndexKey As Variant
    For Each conversationIndexKey In conversationIndexGroups.keys
        Dim groupedMailItems As Collection
        Set groupedMailItems = conversationIndexGroups(conversationIndexKey)
        Set groupedMailItems = SortMailItemsByCreationTime(groupedMailItems) ' 昇順Sort
        
        Dim treeNode As treeNode
        Dim mailItem As mailItem
        Dim firstTreeNode As treeNode
        
        Set firstTreeNode = ProcessGroupedMailItems(groupedMailItems, treeNode, mailItem)
        
        If Not firstTreeNode Is Nothing Then
            Call BuildConversationTree(mailItem, treeNode, originalConversationId)
            Call parentTreeNode.AddChild(firstTreeNode)
        End If
    Next
End Sub

''' conversationIndexでグループ済みのコレクションに対してTreeNodeを構築。自分も宛先に含めて送信している場合に対応するためのもの。ByRefになっているので注意
'''
Private Function ProcessGroupedMailItems(ByVal groupedMailItems As Collection, ByRef processedTreeNode As treeNode, ByRef processedMailItem As mailItem) As treeNode
    Dim rootTreeNode As treeNode
    
    If groupedMailItems.count = 1 Or groupedMailItems.count = 2 Then
        Set rootTreeNode = New treeNode
        Set rootTreeNode.Data = groupedMailItems(1)
    Else
        ' エラーハンドリング（必要に応じて実装）
        Call Err.Raise(1111, , "Mailアイテムの数が異常です。処理を終了します")
        Set ProcessGroupedMailItems = Nothing
        Exit Function
    End If
    
    Select Case groupedMailItems.count
        Case 1
            Set processedTreeNode = rootTreeNode
            Set processedMailItem = groupedMailItems(1)
        Case 2
            '自分も宛先に含めて送信している場合
            Dim replyTreeNode As treeNode
            Set replyTreeNode = New treeNode
            Set replyTreeNode.Data = groupedMailItems(2)
            Call rootTreeNode.AddChild(replyTreeNode)
            
            Set processedTreeNode = replyTreeNode
            Set processedMailItem = groupedMailItems(2)
    End Select
    
    Set ProcessGroupedMailItems = rootTreeNode
End Function

''' 指定されたConversationIDと異なるメールアイテムだけを保持する新しいCollectionを返します。
Private Function FilterMailItemsByDifferentConversationId(ByVal mailItems As Collection, ByVal conversationIdToRemove As String) As Collection
    Dim filteredMailItems As New Collection
    Dim i As Long
    
    For i = 1 To mailItems.count
        If mailItems(i).ConversationID <> conversationIdToRemove Then
            filteredMailItems.Add mailItems(i)
        End If
    Next
    
    Set FilterMailItemsByDifferentConversationId = filteredMailItems
End Function

''' 検索対象のフォルダを取得します
Private Function GetSearchFolders() As Collection
    Set GetSearchFolders = Nothing
    
    Dim serachFolders As Collection
    Set serachFolders = FolderFunc.GetDefaultFolderByMail()
    Dim f As Variant
    For Each f In Split(SEARCH_FOLDER_PATH_LIST)
        Dim folderPath As String
        folderPath = f
        folderPath = TrimWhitespaceAndNewlines(folderPath)
    
        Call serachFolders.Add(FolderFunc.GetFolder(folderPath))
    Next
    Set GetSearchFolders = serachFolders
End Function
