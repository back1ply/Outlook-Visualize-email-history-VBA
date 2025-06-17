'返信元のID
Const PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001e"
'インタネットメッセージID
Const PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001e"
'会話ID
Const PR_CONVERSATION_ID = "http://schemas.microsoft.com/mapi/proptag/0x30130102"

Const PR_PARENT_ENTRY_ID = "http://schemas.microsoft.com/mapi/proptag/0x0e090102"

''' 引数に渡されたobjectがMailItemかを判断します。
''' item: mail / appointment ・・・
''' return: true=maiItem, false=mailItem以外
Public Function IsMailItem(ByVal item As Object) As Boolean
    IsMailItem = False
    If TypeOf item Is mailItem Then
        IsMailItem = True
    End If
End Function

''' itemsからMailItemだけ抜き出して返します。
''' item: SimpleItems
''' return: 抜き出したMailItemコレクション
Public Function GetMailItemOnly(ByVal items As SimpleItems) As Collection
    Dim item As Variant
    Dim mailItems As New Collection
    
    ' SimpleItems内をループ
    For Each item In items
        If IsMailItem(item) Then
            mailItems.Add item
        End If
    Next item
    
    ' Collectionを返す
    Set GetMailItemOnly = mailItems
End Function

''' mailItemCollectionをConvesationIndexでグループ化します。
''' return: グループ化したディクショナリ
Public Function GroupMailItemsByConversationIndex(ByVal mailItemsCollection As Collection) As Object
    Dim dict As Object
    Set dict = GetDictionaryInstance()

    Dim mailItem As mailItem
    Dim convIndex As String
    
    ' コレクション内の各MailItemを処理
    For Each mailItem In mailItemsCollection
        convIndex = mailItem.ConversationIndex
        
        ' ディクショナリにConversationIndexをキーとして追加
        If dict.Exists(convIndex) Then
            ' 既に存在するキーの場合は、そのコレクションに追加

            Call dict(convIndex).Add(mailItem)

        Else
            ' 新しいキーの場合は、新たにコレクションを作成し、追加
            Dim newCollection As Collection
            Set newCollection = New Collection
            Call newCollection.Add(mailItem)
            Call dict.Add(convIndex, newCollection)
        End If
    Next mailItem
   
    ' ディクショナリを返す
    Set GroupMailItemsByConversationIndex = dict
End Function

''' mailItemCollectionをConvesationIdでグループ化します。
''' return: グループ化したディクショナリ
Public Function GroupMailItemsByConversationId(ByVal mailItemsCollection As Collection) As Object
    Dim dict As Object
    Set dict = GetDictionaryInstance()

    Dim mailItem As mailItem
    Dim key As String
    
    ' コレクション内の各MailItemを処理
    For Each mailItem In mailItemsCollection
        key = mailItem.ConversationID
        
        ' ディクショナリにConversationIndexをキーとして追加
        If dict.Exists(key) Then
            ' 既に存在するキーの場合は、そのコレクションに追加

            Call dict(key).Add(mailItem)

        Else
            ' 新しいキーの場合は、新たにコレクションを作成し、追加
            Dim newCollection As Collection
            Set newCollection = New Collection
            Call newCollection.Add(mailItem)
            Call dict.Add(key, newCollection)
        End If
    Next mailItem
   
    ' ディクショナリを返す
    Set GroupMailItemsByConversationId = dict
End Function

''' 現在のmailItemを取得します。
''' return: 取得できたmailItem, 取得できんかった場合はNothingになります。
Public Function GetCurrentMailItem() As mailItem
    Set GetCurrentMailItem = Nothing
    
    Dim item As Object
    Set item = GetCurrentItem()
    
    If IsMailItem(item) Then
        Set GetCurrentMailItem = item
    End If
End Function

''' creationtTimeで昇順にソート
''' return: creationtTimeを昇順
Public Function SortMailItemsByCreationTime(mailItems As Collection) As Collection
    Dim sortedItems As New Collection
    Dim i As Long, j As Long
    Dim tempMailItem As mailItem
    Dim tempIndex As Long

    ' 挿入ソートアルゴリズムを使用して、mailItemsをCreationTimeで昇順にソート
    For i = 1 To mailItems.count
        Set tempMailItem = mailItems(i)
        
        If sortedItems.count = 0 Then
            sortedItems.Add tempMailItem
        Else
            For j = 1 To sortedItems.count
                If tempMailItem.CreationTime < sortedItems(j).CreationTime Then
                    sortedItems.Add tempMailItem, , j
                    Exit For
                End If
            Next j
            
            ' 新しいアイテムが最後までソートされたリストに追加されなかった場合、最後に追加
            If j > sortedItems.count Then
                sortedItems.Add tempMailItem
            End If
        End If
    Next i
    
    Set SortMailItemsByCreationTime = sortedItems
End Function

''' InternetMessageIdをIn-Reply-Toに持つMailItemコレクションを返します
''' id: InternetMessageId
''' targetFolder: 検索するフォルダ
''' return: MailItemコレクション
Public Function GetReplyMailItems(ByVal id As String, ByVal targetFolder As Outlook.folder) As Collection
    Dim filterString As String
    Dim findItems As items
    Dim result As New Collection
    Dim item As Variant
    
    filterString = "@SQL=""" & PR_IN_REPLY_TO_ID & """ = '" & id & "'"
    Set findItems = targetFolder.items.Restrict(filterString)
    
    For Each item In findItems
        Call result.Add(item)
    Next

    Set GetReplyMailItems = result
End Function

''' InternetMessageIdをIn-Reply-Toに持つMailItemを複数のフォルダから探し見つかったMailItemコレクションを返します
''' id: InternetMessageId
''' targetFolder: 検索するフォルダコレクション
''' return: MailItemコレクション
Public Function GetReplyMailItemsInFolders(ByVal id As String, ByVal targetFolders As Collection) As Collection

    Dim replyMailItems As Collection
    Set replyMailItems = New Collection
    
    Dim targetFolder
    For Each targetFolder In targetFolders
        Dim findReplyItems As Collection
        Set findReplyItems = GetReplyMailItems(id, targetFolder)
        
        Dim findReplyItem
        For Each findReplyItem In findReplyItems
            Call replyMailItems.Add(findReplyItem)
        Next
    Next
    
    Set GetReplyMailItemsInFolders = replyMailItems
End Function


''' 引数に渡されたMailItemからインターネットメッセージIDを取得します。
''' item: mailItem
''' return: 取得できたインターネットメッセージID
Public Function GetInternetMessageId(ByVal item As mailItem) As String
    GetInternetMessageId = item.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
End Function

''' 引数に渡されたMailItemから返信元のIDを取得します。
''' item: mailItem
''' return: 取得できた返信元のID
Public Function GetInReplyToId(ByVal item As mailItem) As String
    GetInReplyToId = item.PropertyAccessor.GetProperty(PR_IN_REPLY_TO_ID)
End Function

''' 引数に渡されたMailItemから会話IDをPrpertyAccessorを使用して取得します。
''' item: mailItem
''' return: 取得できた会話ID
Public Function GetConversationIdBySchema(ByVal item As mailItem) As String
    With item.PropertyAccessor
        GetConversationId = .BinaryToString(.GetProperty(PR_CONVERSATION_ID))
    End With
End Function

