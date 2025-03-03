Imports _04AfterDI_Container
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Moq
Imports Unity
Imports Unity.Lifetime

Namespace TestBeforeDI_MockLib
    ' テストコードを Moq を使用して書き換え

    <TestClass>
    Public Class UserServiceTests
        Private _userService As UserService
        Private _mockUserAccess As Mock(Of IUserDataAccess)
        Private _mockFamilyAccess As Mock(Of IFamilyDataAccess)
        Private _mockDepartmentAccess As Mock(Of IDepartmentDataAccess)

        <TestInitialize>
        Public Sub Setup()
            ' Moqを使用してモックを作成
            _mockUserAccess = New Mock(Of IUserDataAccess)()
            _mockFamilyAccess = New Mock(Of IFamilyDataAccess)()
            _mockDepartmentAccess = New Mock(Of IDepartmentDataAccess)()

            ' テスト用のDIコンテナ設定
            Dim container = New UnityContainer()
            container.RegisterInstance(Of IUserDataAccess)(_mockUserAccess.Object)
            container.RegisterInstance(Of IFamilyDataAccess)(_mockFamilyAccess.Object)
            container.RegisterInstance(Of IDepartmentDataAccess)(_mockDepartmentAccess.Object)

            ' DIコンテナからUserServiceを取得
            _userService = container.Resolve(Of UserService)()

        End Sub

        <TestMethod>
        Public Sub Mcok_ユーザ名を正しく表示する()
            ' Arrange
            Dim expectedName = "山田　太郎 様"

            ' モックの振る舞いを設定
            '   _mockDataAccess.Setup(Sub(m) m.GetUser(1)) 'Functionなら戻り値を設定する必要がある
            _mockUserAccess.SetupGet(Function(m) m.UserId).Returns(1)
            _mockUserAccess.SetupGet(Function(m) m.Name1).Returns("太郎")
            _mockUserAccess.SetupGet(Function(m) m.Name2).Returns("山田")

            ' Act
            Dim actualName = _userService.GetUserNameById(1)

            ' Assert
            Assert.AreEqual(expectedName, actualName)

            ' モックメソッドが呼ばれたことを検証
            _mockUserAccess.Verify(Sub(m) m.GetUser(1), Times.Once)
        End Sub

        <TestMethod>
        Public Sub Mock_ユーザが取得できない場合の表示名()
            ' Arrange
            Dim expectedName = "ユーザ情報が取得できません"

            ' モックの振る舞いを設定
            '   _mockDataAccess.Setup(Sub(m) m.GetUser(999)) 'Functionなら戻り値を設定する必要がある
            _mockUserAccess.SetupGet(Function(m) m.UserId).Returns(999)
            _mockUserAccess.SetupGet(Function(m) m.Name1).Returns("")
            _mockUserAccess.SetupGet(Function(m) m.Name2).Returns("")

            ' Act
            Dim actualName = _userService.GetUserNameById(999)

            ' Assert
            Assert.AreEqual(expectedName, actualName)

            ' モックメソッドが呼ばれたことを検証
            _mockUserAccess.Verify(Sub(m) m.GetUser(999), Times.Once)
        End Sub
    End Class
End Namespace

