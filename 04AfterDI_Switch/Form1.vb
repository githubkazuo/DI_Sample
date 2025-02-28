Imports System.ComponentModel
Imports System.Configuration
Imports Unity
Imports Unity.Lifetime

Public Class Form1
    Private Sub btnUserInfo_Click(sender As Object, e As EventArgs) Handles btnUserInfo.Click
        'ユーザ情報を取得してテキストボックスに表示

        ' DIコンテナから UserService のインスタンスを取得
        Dim userService = ContainerConfig.GetConfiguredContainer().Resolve(Of UserService)()

        ' ユーザー情報を取得して表示
        txtName.Text = userService.GetUserNameById(CInt(txtID.Text))
    End Sub

    Private Sub btnSwitchDb_Click(sender As Object, e As EventArgs) Handles btnSwitchDb.Click
        ' データベースを切り替える
        ContainerConfig.SwitchDatabase()
        MessageBox.Show("データベース接続が切り替わりました。")
    End Sub

End Class

'＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

' DIを使用したユーザ情報の取得といろんな処理のクラス
Public Class UserService
    Private _userDataAccess As IUserDataAccess
    Private _orderDataAccess As IOrderDataAccess
    Private _departmentDataAccess As IDepartmentDataAccess

    ' コンストラクタインジェクション
    Public Sub New(userDataAccess As IUserDataAccess,
                   orderDataAccess As IOrderDataAccess,
                   departmentDataAccess As IDepartmentDataAccess)
        _userDataAccess = userDataAccess
        _orderDataAccess = orderDataAccess
        _departmentDataAccess = departmentDataAccess
    End Sub

    Public Function GetUserNameById(id As Integer) As String
        ' 内部フィールドを使用してデータアクセス
        _userDataAccess.GetUser(id)

        Dim FirstName = _userDataAccess.Name1
        Dim LastName = _userDataAccess.Name2

        ' ユーザ名が空の場合、メッセージを返す
        If String.IsNullOrEmpty(FirstName) AndAlso String.IsNullOrEmpty(LastName) Then
            Return "ユーザ情報が取得できません"
        Else
            ' ユーザ名を編集して返す
            Return String.Format("{0}　{1} 様", LastName, FirstName)
        End If
    End Function

    Public Function GetUserOrder(id As Integer) As String
        ' 内部フィールドを使用してデータアクセス
        _orderDataAccess.GetOrder(id)
        Return String.Format("注文ID: {0}, 日付: {1}, 金額: {2}円",
                           _orderDataAccess.OrderId,
                           _orderDataAccess.OrderDate.ToShortDateString(),
                           _orderDataAccess.Amount)
    End Function

    Public Function GetUserDepartment(id As Integer) As String
        ' ユーザーの所属部署を取得
        _departmentDataAccess.GetDepartment(id)
        Return String.Format("部署: {0}", _departmentDataAccess.DepartmentName)
    End Function
End Class

' DIコンテナの設定クラス
Public Class ContainerConfig
    Private Shared _container As IUnityContainer
    Private Shared _currentDbType As String = "SqlServer" ' デフォルトはSQL Server

    Public Shared Function GetConfiguredContainer() As IUnityContainer
        If _container Is Nothing Then
            _container = New UnityContainer()
            RegisterTypes()
        End If
        Return _container
    End Function

    Public Shared Sub SwitchDatabase()
        ' データベースタイプを切り替え
        If _currentDbType = "SqlServer" Then
            _currentDbType = "Oracle"
        Else
            _currentDbType = "SqlServer"
        End If

        ' コンテナを再構築
        _container = New UnityContainer()
        RegisterTypes()
    End Sub

    Private Shared Sub RegisterTypes()
        ' 現在選択されているDBタイプに基づいて依存関係を登録
        If _currentDbType = "SqlServer" Then
            ' SQL Server実装の登録
            _container.RegisterType(Of IUserDataAccess, SqlDataAccessUser)(
                New ContainerControlledLifetimeManager())
            _container.RegisterType(Of IOrderDataAccess, SqlDataAccessOrder)(
                New ContainerControlledLifetimeManager())
            _container.RegisterType(Of IDepartmentDataAccess, SqlDataAccessDepartment)(
                New ContainerControlledLifetimeManager())
        Else
            ' Oracle実装の登録
            _container.RegisterType(Of IUserDataAccess, OracleDataAccessUser)(
                New ContainerControlledLifetimeManager())
            _container.RegisterType(Of IOrderDataAccess, OracleDataAccessOrder)(
                New ContainerControlledLifetimeManager())
            _container.RegisterType(Of IDepartmentDataAccess, OracleDataAccessDepartment)(
                New ContainerControlledLifetimeManager())
        End If

        ' UserServiceの登録
        _container.RegisterType(Of UserService)()
    End Sub
End Class

' ユーザーデータアクセスのインターフェース
Public Interface IUserDataAccess
    Property UserId As Integer
    Property Name1 As String
    Property Name2 As String
    Sub GetUser(id As Integer)
End Interface

Public Interface IOrderDataAccess
    Property OrderId As Integer
    Property OrderDate As Date
    Property Amount As Decimal
    Sub GetOrder(id As Integer)
End Interface

Public Interface IDepartmentDataAccess
    Property DepartmentId As Integer
    Property DepartmentName As String
    Sub GetDepartment(id As Integer)
End Interface

'＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
' SQL Server実装クラス
'＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

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
        Console.WriteLine("SQL Server からユーザー情報を取得しました")
    End Sub
End Class

Public Class SqlDataAccessOrder
    Implements IOrderDataAccess

    Public Property OrderId As Integer Implements IOrderDataAccess.OrderId
    Public Property OrderDate As Date Implements IOrderDataAccess.OrderDate
    Public Property Amount As Decimal Implements IOrderDataAccess.Amount

    Public Sub GetOrder(id As Integer) Implements IOrderDataAccess.GetOrder
        ' SQL Serverから注文情報取得のロジック
        OrderId = id + 1000
        OrderDate = DateTime.Now.AddDays(-5)
        Amount = 12500
        Console.WriteLine("SQL Server から注文情報を取得しました")
    End Sub
End Class

Public Class SqlDataAccessDepartment
    Implements IDepartmentDataAccess

    Public Property DepartmentId As Integer Implements IDepartmentDataAccess.DepartmentId
    Public Property DepartmentName As String Implements IDepartmentDataAccess.DepartmentName

    Public Sub GetDepartment(id As Integer) Implements IDepartmentDataAccess.GetDepartment
        ' SQL Serverから部署情報取得のロジック
        DepartmentId = id Mod 3 + 1
        Select Case DepartmentId
            Case 1
                DepartmentName = "営業部"
            Case 2
                DepartmentName = "開発部"
            Case Else
                DepartmentName = "総務部"
        End Select
        Console.WriteLine("SQL Server から部署情報を取得しました")
    End Sub
End Class

'＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
' Oracle実装クラス - 新しく追加
'＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

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
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
        'IDが999の場合、名前は空で返す
        If id = 999 Then
            UserId = id
            Name1 = ""
            Name2 = ""
        Else
            UserId = id
            Name1 = "次郎" ' Oracleの場合は別のデータを返すと仮定
            Name2 = "佐藤"
        End If
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
        Console.WriteLine("Oracle からユーザー情報を取得しました")
    End Sub
End Class

Public Class OracleDataAccessOrder
    Implements IOrderDataAccess

    Public Property OrderId As Integer Implements IOrderDataAccess.OrderId
    Public Property OrderDate As Date Implements IOrderDataAccess.OrderDate
    Public Property Amount As Decimal Implements IOrderDataAccess.Amount

    Public Sub GetOrder(id As Integer) Implements IOrderDataAccess.GetOrder
        ' Oracleから注文情報取得のロジック
        OrderId = id + 2000 ' Oracle用の処理
        OrderDate = DateTime.Now.AddDays(-3)
        Amount = 15800
        Console.WriteLine("Oracle から注文情報を取得しました")
    End Sub
End Class

Public Class OracleDataAccessDepartment
    Implements IDepartmentDataAccess

    Public Property DepartmentId As Integer Implements IDepartmentDataAccess.DepartmentId
    Public Property DepartmentName As String Implements IDepartmentDataAccess.DepartmentName

    Public Sub GetDepartment(id As Integer) Implements IDepartmentDataAccess.GetDepartment
        ' Oracleから部署情報取得のロジック
        DepartmentId = id Mod 4 + 1
        Select Case DepartmentId
            Case 1
                DepartmentName = "営業第一部"
            Case 2
                DepartmentName = "システム開発部"
            Case 3
                DepartmentName = "人事部"
            Case Else
                DepartmentName = "経理部"
        End Select
        Console.WriteLine("Oracle から部署情報を取得しました")
    End Sub
End Class

'＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
' 設定ファイルから動的にDB設定を読み込む場合の拡張例
'＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

' より実践的なDBファクトリークラス例
Public Class DatabaseFactory
    Public Shared Function GetUserDataAccess() As IUserDataAccess
        Dim dbType As String = ConfigurationManager.AppSettings("DatabaseType")

        Select Case dbType
            Case "SqlServer"
                Return New SqlDataAccessUser()
            Case "Oracle"
                Return New OracleDataAccessUser()
            Case Else
                Throw New ArgumentException("サポートされていないデータベースタイプです: " & dbType)
        End Select
    End Function

    Public Shared Function GetOrderDataAccess() As IOrderDataAccess
        Dim dbType As String = ConfigurationManager.AppSettings("DatabaseType")

        Select Case dbType
            Case "SqlServer"
                Return New SqlDataAccessOrder()
            Case "Oracle"
                Return New OracleDataAccessOrder()
            Case Else
                Throw New ArgumentException("サポートされていないデータベースタイプです: " & dbType)
        End Select
    End Function

    Public Shared Function GetDepartmentDataAccess() As IDepartmentDataAccess
        Dim dbType As String = ConfigurationManager.AppSettings("DatabaseType")

        Select Case dbType
            Case "SqlServer"
                Return New SqlDataAccessDepartment()
            Case "Oracle"
                Return New OracleDataAccessDepartment()
            Case Else
                Throw New ArgumentException("サポートされていないデータベースタイプです: " & dbType)
        End Select
    End Function
End Class

' 設定ファイルから読み込む場合のDIコンテナ設定例（代替実装）
Public Class ConfigDrivenContainer
    Private Shared _container As IUnityContainer

    Public Shared Function GetConfiguredContainer() As IUnityContainer
        If _container Is Nothing Then
            _container = New UnityContainer()
            RegisterTypes()
        End If
        Return _container
    End Function

    Private Shared Sub RegisterTypes()
        ' 設定ファイルからデータベースタイプを取得
        Dim dbType As String = ConfigurationManager.AppSettings("DatabaseType")

        ' データベースタイプに基づいて実装を登録
        Select Case dbType
            Case "SqlServer"
                _container.RegisterType(Of IUserDataAccess, SqlDataAccessUser)(
                    New ContainerControlledLifetimeManager())
                _container.RegisterType(Of IOrderDataAccess, SqlDataAccessOrder)(
                    New ContainerControlledLifetimeManager())
                _container.RegisterType(Of IDepartmentDataAccess, SqlDataAccessDepartment)(
                    New ContainerControlledLifetimeManager())
            Case "Oracle"
                _container.RegisterType(Of IUserDataAccess, OracleDataAccessUser)(
                    New ContainerControlledLifetimeManager())
                _container.RegisterType(Of IOrderDataAccess, OracleDataAccessOrder)(
                    New ContainerControlledLifetimeManager())
                _container.RegisterType(Of IDepartmentDataAccess, OracleDataAccessDepartment)(
                    New ContainerControlledLifetimeManager())
            Case Else
                Throw New ArgumentException("サポートされていないデータベースタイプです: " & dbType)
        End Select

        ' UserServiceの登録
        _container.RegisterType(Of UserService)()
    End Sub

    Public Shared Sub RefreshContainer()
        ' 設定が変更された場合にコンテナを再構築
        _container = New UnityContainer()
        RegisterTypes()
    End Sub
End Class