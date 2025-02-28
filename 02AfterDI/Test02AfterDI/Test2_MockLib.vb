Imports _02AfterDI
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Moq


Namespace TestBeforeDI_MockLib
    ' テストコードを Moq を使用して書き換え

    <TestClass>
    Public Class UserServiceTests
        Private _userService As UserService
        Private _mockDataAccess As Mock(Of IUserDataAccess)

        <TestInitialize>
        Public Sub Setup()
            ' Moqを使用してモックを作成
            _mockDataAccess = New Mock(Of IUserDataAccess)()
            _userService = New UserService(_mockDataAccess.Object)
        End Sub

        <TestMethod>
        Public Sub Mcok_ユーザ名を正しく表示する()
            ' Arrange
            Dim expectedName = "山田　太郎 様"

            ' モックの振る舞いを設定
            '   _mockDataAccess.Setup(Sub(m) m.GetUser(1)) 'Functionなら戻り値を設定する必要がある
            _mockDataAccess.SetupGet(Function(m) m.UserId).Returns(1)
            _mockDataAccess.SetupGet(Function(m) m.Name1).Returns("太郎")
            _mockDataAccess.SetupGet(Function(m) m.Name2).Returns("山田")

            ' Act
            Dim actualName = _userService.GetUserNameById(1)

            ' Assert
            Assert.AreEqual(expectedName, actualName)

            ' モックメソッドが呼ばれたことを検証
            _mockDataAccess.Verify(Sub(m) m.GetUser(1), Times.Once)
        End Sub

        <TestMethod>
        Public Sub Mock_ユーザが取得できない場合の表示名()
            ' Arrange
            Dim expectedName = "ユーザ情報が取得できません"

            ' モックの振る舞いを設定
            '   _mockDataAccess.Setup(Sub(m) m.GetUser(999)) 'Functionなら戻り値を設定する必要がある
            _mockDataAccess.SetupGet(Function(m) m.UserId).Returns(999)
            _mockDataAccess.SetupGet(Function(m) m.Name1).Returns("")
            _mockDataAccess.SetupGet(Function(m) m.Name2).Returns("")

            ' Act
            Dim actualName = _userService.GetUserNameById(999)

            ' Assert
            Assert.AreEqual(expectedName, actualName)

            ' モックメソッドが呼ばれたことを検証
            _mockDataAccess.Verify(Sub(m) m.GetUser(999), Times.Once)
        End Sub
    End Class
End Namespace

