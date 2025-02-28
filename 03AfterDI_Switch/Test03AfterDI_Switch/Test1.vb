Imports _02AfterDI
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace TestBeforeDI
    ' 修正したテストクラス
    <TestClass>
    Public Class UserServiceTests
        Private _userService As UserService
        Private _mockDataAccess As IUserDataAccess

        <TestInitialize>
        Public Sub Setup()
            ' テスト用のモックを使用
            _mockDataAccess = New MockUserDataAccess()
            _userService = New UserService(_mockDataAccess)
        End Sub

        <TestMethod>
        Public Sub ユーザ名を正しく表示する()
            ' Arrange
            Dim expectedName = "山田　太郎 様"

            ' Act
            Dim actualName = _userService.GetUserNameById(1)

            ' Assert
            Assert.AreEqual(expectedName, actualName)
        End Sub

        <TestMethod>
        Public Sub ユーザが取得できない場合の表示名()
            ' Arrange
            Dim expectedName = "ユーザ情報が取得できません"

            ' Act
            Dim actualName = _userService.GetUserNameById(999)

            ' Assert
            Assert.AreEqual(expectedName, actualName)
        End Sub
    End Class

    ' テスト用のモッククラス
    Public Class MockUserDataAccess
        Implements IUserDataAccess

        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
        '＊＊＊DBがなくてもテストが実行できるような仮のデータアクセス処理＊＊＊
        '＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊


        Public Property UserId As Integer Implements IUserDataAccess.UserId
        Public Property Name1 As String Implements IUserDataAccess.Name1
        Public Property Name2 As String Implements IUserDataAccess.Name2

        Public Sub GetUser(id As Integer) Implements IUserDataAccess.GetUser
            ' テスト用のデータを設定
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

End Namespace

