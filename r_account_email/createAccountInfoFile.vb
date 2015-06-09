''' <summary>
''' SCPファイル共有サービスの移行に伴い、新サーバを利用してもらうための
''' アカウント情報をメールで送信するユーティリティである。
''' 
''' 機能
''' １．移行対象のアカウントをアカウント台帳よりロードする（Excel）
''' ２．アカウント設定書毎を作成する。この設定書は新サーバのアカウントとパスワードを記している（Excel）。このファイルにもパスワードを設定する
''' ３．アカウント毎にメールを送信する。メールには以下の添付ファイルを含める
''' 　　・新サーバ利用するための手順書
''' 　　・アカウント設定書
''' ４．アカウント毎にアカウント設定書のパスワードを記したメールを送信する
''' 
''' 機能より必要な外部ファイル。ならいびに外部ファイルに必要な情報
''' ・アカウント台帳
''' ・アカウント設定書のテンプレート
''' ・
''' </summary>
''' <remarks></remarks>
Module createAccountInfoFile

    Public Sub main()

        Dim _d2e As jp.polestar.io.dt2excel = New jp.polestar.io.dt2excel
        Dim _d2t As jp.polestar.io.Datatable2TSV = New jp.polestar.io.Datatable2TSV
        Dim _dt As System.Data.DataTable = _d2e.ToDataTable("master\アカウント台帳.xls", "Sheet1", 25, 5)
        Dim _sendList As System.Data.DataTable = createAccountInfoFile(_dt, "template\アカウント設定書テンプレート.xls", "Attachment", "hogehoge")

        _d2t.dt2tsv(_sendList, "sendlist.txt")

        For Each row As System.Data.DataRow In _sendList.Rows
            If row(2).ToString.Length > 0 And row(3).ToString.Length > 0 And row(4).ToString.Length > 0 Then

                Call sendAccountInfoMail(row(2), row(4), row(3))
                Call sendPasswordMail(row(2), row(3))
                ' wait
                System.Threading.Thread.Sleep(1000)
                Console.WriteLine("send email [" & row(2).ToString & "]" & vbTab & row(3).ToString)
            Else
                Console.WriteLine("send skip email [" & row(2).ToString & "]" & vbTab & row(3).ToString)

            End If

        Next

    End Sub

    Public Sub sendAccountInfoMail(toName As String, attacheFile As String, email As String)

        ' https://support.microsoft.com/ja-jp/kb/313803/ja

        ' Create an Outlook application.
        Dim _outlook As Microsoft.Office.Interop.Outlook.Application
        _outlook = New Microsoft.Office.Interop.Outlook.Application

        ' Create a new MailItem.
        Dim _Message As Microsoft.Office.Interop.Outlook.MailItem
        _Message = _outlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
        _Message.Subject = "【HOGE】アカウント発送のご案内"

        ' TODO: Replace with a valid e-mail address.
        _Message.To = email
        _Message.CC = ""

        Dim _rst As System.IO.StreamReader = New System.IO.StreamReader("template\MailBody01.txt", System.Text.Encoding.UTF8)
        _Message.Body = _rst.ReadToEnd
        _Message.Body = _Message.Body.Replace("___NAME___", toName)
        _Message.Body = _Message.Body & vbCrLf & vbCrLf

        Dim _configFile As System.IO.FileInfo = New System.IO.FileInfo(attacheFile)
        Dim FileSource_1 As String = _configFile.FullName
        Dim FileDisp_1 As String = _configFile.Name

        Dim _manfile = New System.IO.FileInfo("master\Cworks物件管理.pdf")
        Dim FileSource_2 As String = _manfile.FullName
        Dim FileDisp_2 As String = _manfile.Name

        _Message.Attachments.Add(FileSource_1, , , FileDisp_1)
        _Message.Attachments.Add(FileSource_2, , , FileDisp_2)

        ' Send
        _Message.Send()

        ' Clean up
        _outlook = Nothing
        _Message = Nothing

    End Sub

    Public Sub sendPasswordMail(toName As String, email As String)


        ' https://support.microsoft.com/ja-jp/kb/313803/ja

        ' Create an Outlook application.
        Dim _outlook As Microsoft.Office.Interop.Outlook.Application
        _outlook = New Microsoft.Office.Interop.Outlook.Application

        ' Create a new MailItem.
        Dim _Message As Microsoft.Office.Interop.Outlook.MailItem
        _Message = _outlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
        _Message.Subject = "【HOGE】パスワードのご案内"

        ' TODO: Replace with a valid e-mail address.
        _Message.To = email
        _Message.CC = ""

        Dim _rst As System.IO.StreamReader = New System.IO.StreamReader("template\MailBody02.txt", System.Text.Encoding.UTF8)
        _Message.Body = _rst.ReadToEnd
        _Message.Body = _Message.Body.Replace("___NAME___", toName)
        _Message.Body = _Message.Body & vbCrLf & vbCrLf

        ' Send
        _Message.Send()

        ' Clean up
        _outlook = Nothing
        _Message = Nothing



    End Sub


    Public Function createAccountInfoFile(_dt As System.Data.DataTable, templateFile As String, outputDirectory As String, password As String) As System.Data.DataTable
        Dim result As System.Data.DataTable = New System.Data.DataTable("SEND_LIST")

        result.Columns.Add("NO", GetType(Integer))
        result.Columns.Add("ID", GetType(String))
        result.Columns.Add("NAME", GetType(String))
        result.Columns.Add("EMAIL", GetType(String))
        result.Columns.Add("ATTACH_FILE", GetType(String))

        For Each _row As System.Data.DataRow In _dt.Rows

            Dim _wbs As SpreadsheetGear.IWorkbookSet = SpreadsheetGear.Factory.GetWorkbookSet(System.Globalization.CultureInfo.CurrentCulture)
            Dim _wb As SpreadsheetGear.IWorkbook = _wbs.Workbooks.Open(templateFile)
            Dim fileName As String = outputDirectory & "\【IDP】" & _row(2) & " アカウント設定書.xlsx"
            Dim row As System.Data.DataRow = result.NewRow

            Try
                row("NO") = _row(0)     ' NO

                _wb.Worksheets(0).Cells(1, 0).Value = _row(1)   '　ID
                row("ID") = _row(1)

                _wb.Worksheets(0).Cells(1, 1).Value = _row(2)   '　氏名
                row("NAME") = _row(2)

                _wb.Worksheets(0).Cells(1, 2).Value = _row(3)   '   ひらがな
                _wb.Worksheets(0).Cells(1, 3).Value = _row(4)   '   会社名
                _wb.Worksheets(0).Cells(1, 4).Value = _row(5)   '   部署

                row("EMAIL") = _row(6) ' EMail

                _wb.Worksheets(0).Cells(1, 5).Value = _row(7)   '   アカウント名
                _wb.Worksheets(0).Cells(1, 6).Value = _row(8)   '   パスワード

                _wb.Worksheets(0).Cells(1, 7).Value = _row(9)   '   権限1
                _wb.Worksheets(0).Cells(1, 8).Value = _row(10)   '   権限2
                _wb.Worksheets(0).Cells(1, 9).Value = _row(11)  '   権限3
                _wb.Worksheets(0).Cells(1, 10).Value = _row(12)   '   権限4
                _wb.Worksheets(0).Cells(1, 11).Value = _row(13)   '   権限5
                _wb.Worksheets(0).Cells(1, 12).Value = _row(14)   '   権限6
                _wb.Worksheets(0).Cells(1, 13).Value = _row(15)   '   権限7
                _wb.Worksheets(0).Cells(1, 14).Value = _row(16)   '   権限8
                _wb.Worksheets(0).Cells(1, 15).Value = _row(17)   '   権限9
                _wb.Worksheets(0).Cells(1, 16).Value = _row(18)   '   権限10
                _wb.Worksheets(0).Cells(1, 17).Value = _row(19)   '   権限11
                _wb.Worksheets(0).Cells(1, 18).Value = _row(20)   '   権限12
                _wb.Worksheets(0).Cells(1, 19).Value = _row(21)   '   権限13
                _wb.Worksheets(0).Cells(1, 20).Value = _row(22)   '   権限14
                _wb.Worksheets(0).Cells(1, 21).Value = _row(23)   '   権限15
                _wb.Worksheets(0).Cells(1, 22).Value = _row(24)   '   権限16

                _wb.SaveAs(fileName, SpreadsheetGear.FileFormat.OpenXMLWorkbook, password)

                row("ATTACH_FILE") = fileName
                _wb.Close()

                Console.WriteLine("create attach file " & _row(1) & vbTab & _row(2))


            Catch ex As Exception

                row("ATTACH_FILE") = String.Empty
                Console.WriteLine("*********************************************************************")
                Console.WriteLine("exception attach file " & _row(1) & vbTab & _row(2))
                Console.WriteLine(ex.ToString)
            Finally
                _wb = Nothing
                _wbs = Nothing

            End Try

            result.Rows.Add(row)
        Next

        Return result

    End Function

End Module
