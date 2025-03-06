Public Class Form1
    Private Sub btnUserInfo_Click(sender As Object, e As EventArgs) Handles btnUserInfo.Click
        'ユーザ情報を取得してテキストボックスに表示

        ' 依存性の注入（SQL Serverのユーザ情報を取得するクラスを
        Dim dataAccess = New SqlDataAccessUser()
        Dim userService = New UserService(dataAccess)

        userService.GetUserNameById(CInt(txtID.Text))
        txtName.Text = userService.GetUserNameById(CInt(txtID.Text))
    End Sub
End Class

'＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊



' DIを使用したユーザ情報の取得といろんな処理のクラス
Public Class UserService

    ' データアクセス部品
    Private ReadOnly _dataAccess As IUserDataAccess

    ' コンストラクタで依存性を注入（インターフェースで受ける）
    Public Sub New(dataAccess As IUserDataAccess)
        _dataAccess = dataAccess
    End Sub

    Public Function GetUserNameById(id As Integer) As String
        '外部から受け取ったデータアクセス部品（依存性）でデータ取得
        _dataAccess.GetUser(id)

        Dim FirstName = _dataAccess.Name1
        Dim LastName = _dataAccess.Name2

        'ユーザ名が空の場合、メッセージを返す
        If String.IsNullOrEmpty(FirstName) AndAlso String.IsNullOrEmpty(LastName) Then
            Return "ユーザ情報が取得できません"
        Else
            ' ユーザ名を編集して返す
            Return String.Format("{0}　{1} 様", LastName, FirstName)
        End If
    End Function

End Class


' ユーザーデータアクセスのインターフェース
Public Interface IUserDataAccess
    Property UserId As Integer
    Property Name1 As String
    Property Name2 As String
    Sub GetUser(id As Integer)
End Interface


' 仮想のSQL Serverのユーザ情報に関するデータアクセスクラス（依存性に該当）
' 既存のSqlDataAccessUserクラスをインターフェースの実装として修正
Public Class SqlDataAccessUser
    Implements IUserDataAccess

    Public Property UserId As Integer Implements IUserDataAccess.UserId
    Public Property Name1 As String Implements IUserDataAccess.Name1
    Public Property Name2 As String Implements IUserDataAccess.Name2

    Public Sub GetUser(id As Integer) Implements IUserDataAccess.GetUser

        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
        '＊＊＊SQL Serverからユーザ情報を取得する処理＊＊＊
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

        '（仮想の処理）
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
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
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
    End Sub
End Class




