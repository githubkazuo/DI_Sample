Imports BeforeDI
Imports BeforeDI.Form1
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace TestBeforeDI

    <TestClass>
    Public Class UserServiceTests
        Private _userService As UserService

        <TestInitialize>
        Public Sub Setup()
            _userService = New UserService()
        End Sub

        <TestMethod>
        Public Sub ユーザ名を正しく表示する()
            ' Arrange
            Dim expectedName = "山田　太郎 様" ' SqlDataAccessの固定値に依存

            ' Act
            Dim actualName = _userService.GetUserNameById(1)　'"山田　太郎 様"を返してくるDBを用意しておく必要がある。

            ' Assert
            Assert.AreEqual(expectedName, actualName)
        End Sub

        <TestMethod>
        Public Sub ユーザが取得できない場合の表示名()
            ' Arrange
            Dim expectedName = "ユーザ情報が取得できません" ' SqlDataAccessの固定値に依存

            ' Act
            Dim actualName = _userService.GetUserNameById(999)　'999では取得できないDBを用意しておく必要がある。

            ' Assert
            Assert.AreEqual(actualName, expectedName)
        End Sub

    End Class
End Namespace

