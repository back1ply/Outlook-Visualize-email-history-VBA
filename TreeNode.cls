Private m_Parent As treeNode
Private m_Children As Collection ' TreeNodeClass
Public Data As Variant

Private Sub Class_Initialize()
    Set m_Children = New Collection
End Sub

'親ノードのセット
Public Property Set Parent(ByVal value As treeNode)
    Set m_Parent = value
End Property

'親ノードの取得
Public Property Get Parent() As treeNode
    Set Parent = m_Parent
End Property

' 子ノードのコレクションの取得
Public Property Get Children() As Collection
    Set Children = m_Children
End Property

'''' メソッド
''' ####################################################################################################################

' 子ノードを追加するメソッド
Public Sub AddChild(ByVal childNode As treeNode)
    Set childNode.Parent = Me
    Call m_Children.Add(childNode)
End Sub

Public Function GetAllChildrenCount() As Long
    Dim totalCount As Long
    totalCount = 0
    
    ' 再帰的に全子ノードのカウントを行う
    Dim childNode As treeNode
    For Each childNode In m_Children
        totalCount = totalCount + 1 ' 自分自身をカウント
        totalCount = totalCount + childNode.GetAllChildrenCount ' 子ノードのカウントを加算
    Next
    
    GetAllChildrenCount = totalCount
End Function

