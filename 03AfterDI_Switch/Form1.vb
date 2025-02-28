Imports System.Data.Odbc

Public Class Form1

    Private Sub btnUserInfo_Click(sender As Object, e As EventArgs) Handles btnUserInfo.Click
        ' ファクトリークラスを使って現在のデータベース実装を取得
        Dim dataAccess = DatabaseFactory.GetUserDataAccess()
        Dim userService = New UserService()

        txtName.Text = userService.GetUserNameById(CInt(txtID.Text), dataAccess)
    End Sub

    Private Sub btnSwitchDb_Click(sender As Object, e As EventArgs) Handles btnSwitchDb.Click
        ' ファクトリークラスでデータベースを切り替え
        DatabaseFactory.SwitchDatabase()

        lblCurrentDb.Text = "現在のDB: " & DatabaseFactory.CurrentDatabaseName
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblCurrentDb.Text = "現在のDB: " & DatabaseFactory.CurrentDatabaseName
    End Sub
End Class

'＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

' ユーザーデータアクセスのインターフェース
Public Interface IUserDataAccess
    Property UserId As Integer
    Property Name1 As String
    Property Name2 As String
    Sub GetUser(id As Integer)
End Interface

' DIを使用したユーザ情報の取得といろんな処理のクラス
Public Class UserService

    Public Function GetUserNameById(id As Integer, dataAccess As IUserDataAccess) As String
        '外部から受け取ったデータアクセス部品（依存性）でデータ取得
        dataAccess.GetUser(id)

        Dim FirstName = dataAccess.Name1
        Dim LastName = dataAccess.Name2

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
    Implements IUserDataAccess

    Public Property UserId As Integer Implements IUserDataAccess.UserId
    Public Property Name1 As String Implements IUserDataAccess.Name1
    Public Property Name2 As String Implements IUserDataAccess.Name2

    Public Sub GetUser(id As Integer) Implements IUserDataAccess.GetUser
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
        '＊＊＊SQL Serverからユーザ情報を取得する処理＊＊＊
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

        '（仮想の処理）
        'IDが999の場合、名前は空で返す
        If id = 999 Then
            UserId = id
            Name1 = ""
            Name2 = ""
        Else
            UserId = id
            Name1 = "太郎" ' SQL Server固有のデータ
            Name2 = "山田"
        End If
    End Sub
End Class

' 新規追加: Oracle用のデータアクセスクラス
Public Class OracleDataAccessUser
    Implements IUserDataAccess

    Public Property UserId As Integer Implements IUserDataAccess.UserId
    Public Property Name1 As String Implements IUserDataAccess.Name1
    Public Property Name2 As String Implements IUserDataAccess.Name2

    Public Sub GetUser(id As Integer) Implements IUserDataAccess.GetUser
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
        '＊＊＊Oracleからユーザ情報を取得する処理＊＊＊
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

        '（仮想の処理）
        'IDが999の場合、名前は空で返す
        If id = 999 Then
            UserId = id
            Name1 = ""
            Name2 = ""
        Else
            UserId = id
            Name1 = "太郎（Oracleから）" ' Oracle固有のデータ
            Name2 = "山田"
        End If
    End Sub
End Class

' ファクトリークラス（オプション）: より拡張性の高い実装
Public Class DatabaseFactory
    Private Shared dbType As String = "SQL" '通常は環境変数や、設定ファイルで定義

    Public Shared Function GetUserDataAccess() As IUserDataAccess
        If dbType = "Oracle" Then
            Return New OracleDataAccessUser()
        ElseIf dbType = "SQL" Then
            Return New SqlDataAccessUser()
        Else
            Return New Exception("データベースの種類が不正です")
        End If
    End Function

    Public Shared Sub SwitchDatabase()
        If dbType = "Oracle" Then
            dbType = "SQL"
        ElseIf dbType = "SQL" Then
            dbType = "Oracle"
        End If
    End Sub

    Public Shared ReadOnly Property CurrentDatabaseName As String
        Get
            Return dbType
        End Get
    End Property
End Class

