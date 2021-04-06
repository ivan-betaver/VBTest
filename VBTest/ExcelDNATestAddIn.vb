Imports System.Windows.Forms
Imports ExcelDna.Integration
Imports ExcelDna.IntelliSense

Public Class ExcelDNATestAddIn
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        IntelliSenseServer.Install()
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        IntelliSenseServer.Uninstall()
    End Sub

    <ExcelFunction("VB.NET test function")>
    Public Shared Function addSum(<ExcelArgument("Введите значение")> input As Object) As Object
        MessageBox.Show(input.ToString)
        Return input
    End Function

End Class
