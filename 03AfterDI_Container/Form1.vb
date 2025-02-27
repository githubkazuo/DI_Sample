Imports System.ComponentModel
Imports Unity
Imports Unity.Lifetime

Public Class Form1
    Private Sub btnUserInfo_Click(sender As Object, e As EventArgs) Handles btnUserInfo.Click
        'ユーザ情報を取得してテキストボックスに表示

        '依存性が増えると引数も増える
        Dim userService_1 = New UserService(New SqlDataAccessUser(), New SqlDataAccessOrder(), New SqlDataAccessDepartment())

        ' DIコンテナから UserService のインスタンスを取得
        Dim userService_2 = ContainerConfig.GetConfiguredContainer().Resolve(Of UserService)()

        userService_2.GetUserOrder(CInt(txtID.Text))
        txtName.Text = userService_2.GetUserOrder(CInt(txtID.Text))
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
        ' _orderDataAccessや_departmentDataAccessを使った処理

        Return "hogehoge"
    End Function


End Class


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
        ' 各インターフェースと実装クラスの登録
        _container.RegisterType(Of IUserDataAccess, SqlDataAccessUser)(
            New ContainerControlledLifetimeManager())
        _container.RegisterType(Of IOrderDataAccess, SqlDataAccessOrder)(
            New ContainerControlledLifetimeManager())
        _container.RegisterType(Of IDepartmentDataAccess, SqlDataAccessDepartment)(
            New ContainerControlledLifetimeManager())

        ' UserServiceの登録
        _container.RegisterType(Of UserService)()
    End Sub
End Class


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

Public Class SqlDataAccessOrder
    Implements IOrderDataAccess

    Public Property OrderId As Integer Implements IOrderDataAccess.OrderId
    Public Property OrderDate As Date Implements IOrderDataAccess.OrderDate
    Public Property Amount As Decimal Implements IOrderDataAccess.Amount

    Public Sub GetOrder(id As Integer) Implements IOrderDataAccess.GetOrder
        ' 注文情報取得のロジック
    End Sub
End Class

Public Class SqlDataAccessDepartment
    Implements IDepartmentDataAccess

    Public Property DepartmentId As Integer Implements IDepartmentDataAccess.DepartmentId
    Public Property DepartmentName As String Implements IDepartmentDataAccess.DepartmentName

    Public Sub GetDepartment(id As Integer) Implements IDepartmentDataAccess.GetDepartment
        ' 部署情報取得のロジック
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