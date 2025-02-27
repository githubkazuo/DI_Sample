Public Class Form1
    Private Sub btnUserInfo_Click(sender As Object, e As EventArgs) Handles btnUserInfo.Click
        'ユーザ情報を取得してテキストボックスに表示
        Dim userService = New UserService()
        txtName.Text = userService.GetUserNameById(CInt(txtID.Text))
    End Sub
End Class

'＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊



' DIを使用しないユーザ情報処理クラス
Public Class UserService

    Public Function GetUserNameById(id As Integer) As String
        ' データアクセス部品（依存性）を直接newしている
        Dim _sqlDataAccess As New SqlDataAccessUser()
        _sqlDataAccess.GetUser(id)

        Dim FirstName = _sqlDataAccess.Name1
        Dim LastName = _sqlDataAccess.Name2

        'ユーザ名が空の場合、メッセージを返す
        If String.IsNullOrEmpty(FirstName) AndAlso String.IsNullOrEmpty(LastName) Then
            Return "ユーザ情報が取得できません"
        Else
            ' ユーザ名を編集して返す
            Return String.Format("{0}　{1} 様", LastName, FirstName)
        End If
    End Function

End Class



' 仮想のSQL Serverのユーザ情報に関するデータアクセスクラス（依存性に該当）
Public Class SqlDataAccessUser

    'テーブルのカラム名にあわせてプロパティを定義
    Public Property UserId As Integer
    Public Property Name1 As String
    Public Property Name2 As String

    Public Sub GetUser(id As Integer)

        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
        '＊＊＊SQL Serverからユーザ情報を取得する処理＊＊＊
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

        '（以下は仮想の処理）
        'IDが999の場合、名前は空で返す
        If id = 999 Then
            UserId = id
            Name1 = ""
            Name2 = ""
        Else
            UserId = id
            Name1 = "太郎"
            Name2 = "山田"
        End If
    End Sub
End Class

