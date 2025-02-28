Imports System.ComponentModel
Imports Unity
Imports Unity.Lifetime

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblCurrentDb.Text = "現在のDB: " & DatabaseFactory.CurrentDatabaseName
    End Sub

    Private Sub btnUserInfo_Click(sender As Object, e As EventArgs) Handles btnUserInfo.Click
        '（普通にインスタンスを生成する場合）
        Dim userService_2 = New UserService(DatabaseFactory.GetUserDataAccess(), DatabaseFactory.GetFamilyDataAccess(), DatabaseFactory.GetDepartmentDataAccess())

        ' DIコンテナから UserService のインスタンスを取得
        Dim userService = ContainerConfig.GetConfiguredContainer().Resolve(Of UserService)()

        txtName.Text = userService.GetUserNameById(CInt(txtID.Text))
    End Sub

    Private Sub btnSwitchDb_Click(sender As Object, e As EventArgs) Handles btnSwitchDb.Click
        ' ファクトリークラスでデータベースを切り替え
        DatabaseFactory.SwitchDatabase()

        ' 依存性の再登録
        ContainerConfig.RefreshContainer()

        lblCurrentDb.Text = "現在のDB: " & DatabaseFactory.CurrentDatabaseName
    End Sub
End Class

'＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊




' DIコンテナの設定クラス
Public Class ContainerConfig
    Private Shared _container As IUnityContainer

    Public Shared Function GetConfiguredContainer() As IUnityContainer
        If _container Is Nothing Then
            _container = New UnityContainer()
            RegisterTypes()
        End If
        Return _container
    End Function

    Private Shared Sub RegisterTypes()
        ' 既存の登録をクリア
        _container = New UnityContainer()

        ' データベースの種類に応じた依存性の登録
        If DatabaseFactory.CurrentDatabaseName = "Oracle" Then
            ' Oracle用の実装を登録
            _container.RegisterType(Of IUserDataAccess, OracleDataAccessUser)(
                New ContainerControlledLifetimeManager())
            _container.RegisterType(Of IFamilyDataAccess, OracleFamilyDataAccess)(
                New ContainerControlledLifetimeManager())
            _container.RegisterType(Of IDepartmentDataAccess, OracleDepartmentDataAccess)(
                New ContainerControlledLifetimeManager())
        Else
            ' SQL Server用の実装を登録
            _container.RegisterType(Of IUserDataAccess, SqlDataAccessUser)(
                New ContainerControlledLifetimeManager())
            _container.RegisterType(Of IFamilyDataAccess, SqlFamilyDataAccess)(
                New ContainerControlledLifetimeManager())
            _container.RegisterType(Of IDepartmentDataAccess, SqlDepartmentDataAccess)(
                New ContainerControlledLifetimeManager())
        End If

        ' UserServiceの登録（依存性が自動的に解決される）
        _container.RegisterType(Of UserService)()
    End Sub

    ' データベース切り替え時に依存性を再登録するメソッド
    Public Shared Sub RefreshContainer()
        RegisterTypes()
    End Sub
End Class


' ユーザーデータアクセスのインターフェース
Public Interface IUserDataAccess
    Property UserId As Integer
    Property Name1 As String
    Property Name2 As String
    Sub GetUser(id As Integer)
End Interface

' 家族データアクセスのインターフェース
Public Interface IFamilyDataAccess
    Sub GetFamilyInfo(userId As Integer)
    Function GetFamilyCount(userId As Integer) As Integer
    Function GetFamilyNames(userId As Integer) As List(Of String)
End Interface

' 部署データアクセスのインターフェース
Public Interface IDepartmentDataAccess
    Sub GetDepartmentInfo(userId As Integer)
    Function GetDepartmentName(userId As Integer) As String
    Function GetDepartmentCode(userId As Integer) As String
End Interface

' DIを使用したユーザ情報の取得といろんな処理のクラス（コンストラクタで依存性を受け取るように修正）
Public Class UserService
    Private ReadOnly _userDataAccess As IUserDataAccess
    Private ReadOnly _familyDataAccess As IFamilyDataAccess
    Private ReadOnly _departmentDataAccess As IDepartmentDataAccess

    ' コンストラクタで依存性を受け取る
    Public Sub New(userDataAccess As IUserDataAccess,
                  familyDataAccess As IFamilyDataAccess,
                  departmentDataAccess As IDepartmentDataAccess)
        _userDataAccess = userDataAccess
        _familyDataAccess = familyDataAccess
        _departmentDataAccess = departmentDataAccess
    End Sub

    Public Function GetUserNameById(id As Integer) As String
        ' 内部プロパティの依存性を利用
        _userDataAccess.GetUser(id)

        Dim FirstName = _userDataAccess.Name1
        Dim LastName = _userDataAccess.Name2

        'ユーザ名が空の場合、メッセージを返す
        If String.IsNullOrEmpty(FirstName) AndAlso String.IsNullOrEmpty(LastName) Then
            Return "ユーザ情報が取得できません"
        Else
            ' ユーザ名を編集して返す
            Return String.Format("{0}　{1} 様", LastName, FirstName)
        End If
    End Function

    ' 家族情報を取得するメソッド
    Public Function GetFamilyInfo(userId As Integer) As String
        Dim familyCount = _familyDataAccess.GetFamilyCount(userId)
        If familyCount > 0 Then
            Return String.Format("家族は {0} 名です", familyCount)
        Else
            Return "家族情報はありません"
        End If
    End Function

    ' 部署情報を取得するメソッド
    Public Function GetDepartmentInfo(userId As Integer) As String
        Dim deptName = _departmentDataAccess.GetDepartmentName(userId)
        Dim deptCode = _departmentDataAccess.GetDepartmentCode(userId)

        If String.IsNullOrEmpty(deptName) Then
            Return "部署情報はありません"
        Else
            Return String.Format("部署: {0}（{1}）", deptName, deptCode)
        End If
    End Function
End Class

' 仮想のSQL Serverのユーザ情報に関するデータアクセスクラス
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

' Oracle用のデータアクセスクラス
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

' SQL Server用の家族データアクセスクラス
Public Class SqlFamilyDataAccess
    Implements IFamilyDataAccess

    Public Sub GetFamilyInfo(userId As Integer) Implements IFamilyDataAccess.GetFamilyInfo
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
        '＊＊＊SQL Serverから家族情報を取得する処理＊＊＊
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
    End Sub

    Public Function GetFamilyCount(userId As Integer) As Integer Implements IFamilyDataAccess.GetFamilyCount
        ' 仮の実装: SQL Serverからの家族数を返す
        Return If(userId Mod 2 = 0, 2, 1) ' 偶数IDなら2人、奇数IDなら1人
    End Function

    Public Function GetFamilyNames(userId As Integer) As List(Of String) Implements IFamilyDataAccess.GetFamilyNames
        ' 仮の実装: SQL Serverからの家族名を返す
        Dim names As New List(Of String)
        If userId Mod 2 = 0 Then
            names.Add("山田 花子")
            names.Add("山田 健太")
        Else
            names.Add("山田 花子")
        End If
        Return names
    End Function
End Class

' Oracle用の家族データアクセスクラス
Public Class OracleFamilyDataAccess
    Implements IFamilyDataAccess

    Public Sub GetFamilyInfo(userId As Integer) Implements IFamilyDataAccess.GetFamilyInfo
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
        '＊＊＊Oracleから家族情報を取得する処理＊＊＊
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
    End Sub

    Public Function GetFamilyCount(userId As Integer) As Integer Implements IFamilyDataAccess.GetFamilyCount
        ' 仮の実装: Oracleからの家族数を返す
        Return If(userId Mod 2 = 0, 3, 2) ' 偶数IDなら3人、奇数IDなら2人（Oracle固有のデータ）
    End Function

    Public Function GetFamilyNames(userId As Integer) As List(Of String) Implements IFamilyDataAccess.GetFamilyNames
        ' 仮の実装: Oracleからの家族名を返す
        Dim names As New List(Of String)
        If userId Mod 2 = 0 Then
            names.Add("山田 花子（Oracleから）")
            names.Add("山田 健太（Oracleから）")
            names.Add("山田 裕子（Oracleから）")
        Else
            names.Add("山田 花子（Oracleから）")
            names.Add("山田 健太（Oracleから）")
        End If
        Return names
    End Function
End Class

' SQL Server用の部署データアクセスクラス
Public Class SqlDepartmentDataAccess
    Implements IDepartmentDataAccess

    Public Sub GetDepartmentInfo(userId As Integer) Implements IDepartmentDataAccess.GetDepartmentInfo
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
        '＊＊＊SQL Serverから部署情報を取得する処理＊＊＊
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
    End Sub

    Public Function GetDepartmentName(userId As Integer) As String Implements IDepartmentDataAccess.GetDepartmentName
        ' 仮の実装: SQL Serverからの部署名を返す
        Return If(userId Mod 3 = 0, "営業部", If(userId Mod 3 = 1, "開発部", "総務部"))
    End Function

    Public Function GetDepartmentCode(userId As Integer) As String Implements IDepartmentDataAccess.GetDepartmentCode
        ' 仮の実装: SQL Serverからの部署コードを返す
        Return If(userId Mod 3 = 0, "S001", If(userId Mod 3 = 1, "D001", "G001"))
    End Function
End Class

' Oracle用の部署データアクセスクラス
Public Class OracleDepartmentDataAccess
    Implements IDepartmentDataAccess

    Public Sub GetDepartmentInfo(userId As Integer) Implements IDepartmentDataAccess.GetDepartmentInfo
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
        '＊＊＊Oracleから部署情報を取得する処理＊＊＊
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
    End Sub

    Public Function GetDepartmentName(userId As Integer) As String Implements IDepartmentDataAccess.GetDepartmentName
        ' 仮の実装: Oracleからの部署名を返す
        Return If(userId Mod 3 = 0, "営業部（Oracleから）", If(userId Mod 3 = 1, "開発部（Oracleから）", "総務部（Oracleから）"))
    End Function

    Public Function GetDepartmentCode(userId As Integer) As String Implements IDepartmentDataAccess.GetDepartmentCode
        ' 仮の実装: Oracleからの部署コードを返す
        Return If(userId Mod 3 = 0, "SALES", If(userId Mod 3 = 1, "DEV", "ADMIN"))
    End Function
End Class

' ファクトリークラス: より拡張性の高い実装
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

    Public Shared Function GetFamilyDataAccess() As IFamilyDataAccess
        If dbType = "Oracle" Then
            Return New OracleFamilyDataAccess()
        ElseIf dbType = "SQL" Then
            Return New SqlFamilyDataAccess()
        Else
            Return New Exception("データベースの種類が不正です")
        End If
    End Function

    Public Shared Function GetDepartmentDataAccess() As IDepartmentDataAccess
        If dbType = "Oracle" Then
            Return New OracleDepartmentDataAccess()
        ElseIf dbType = "SQL" Then
            Return New SqlDepartmentDataAccess()
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