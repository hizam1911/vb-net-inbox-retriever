Imports System.Globalization
Imports System.Text
Imports System.IO
Imports S22.Imap
Imports System.Net.Mail

Module Module1
    Dim sBody As String
    Dim sBodyArr As String()
    '-------
    Dim sLine01 As String = ""
    Dim sLine02 As String = ""
    Dim sLine03 As String = ""
    Dim sLine04 As String = ""
    '-------
    Dim sFileName As String = ""
    Dim sFullFolderPath As String = ""
    Dim sEmailContents As String = ""
    '--------

    'Function _generateFileName(ByVal sequence As Integer) As String
    '    Dim currentDateTime As DateTime = DateTime.Now
    '    Return String.Format("{0}-{1:000}-{2:000}.eml",
    '                        currentDateTime.ToString("yyyyMMddHHmmss", New CultureInfo("en-US")),
    '                        currentDateTime.Millisecond,
    '                        sequence)
    'End Function

    Sub Main()

        Try
            ' Gmail IMAP server is "imap.gmail.com"
            Dim oClient As New ImapClient("imap.gmail.com", 993, True)
            Console.WriteLine("Logging in ...")
            oClient.Login("kokojarbot@gmail.com", "AstaYuno123", AuthMethod.Auto)

            Console.WriteLine("Searching for new messages ...")
            Dim uids As IEnumerable(Of UInteger) = oClient.Search(SearchCondition.Unseen())
            Console.WriteLine("Total {0} unread email(s)", uids.Count)
            Dim uidArray As UInteger() = uids.ToArray()

            'get messages
            Dim messages As IEnumerable(Of MailMessage) = oClient.GetMessages(uids, FetchOptions.NoAttachments, True)

            If messages Is Nothing Or messages.Count = 0 Then
                Console.WriteLine("No new email(s)")
            Else
                For Each message In messages
                    sBody = message.Body.ToString
                    sBodyArr = sBody.Split(":")

                    sBodyArr(1) = Replace(sBodyArr(1), "Do you have a benefits administrator", "")
                    sBodyArr(2) = Replace(sBodyArr(2), "Email", "")
                    sBodyArr(3) = Replace(sBodyArr(3), "Mobile Phone Number", "")

                    sBodyArr(1) = Replace(sBodyArr(1), vbCrLf, "")
                    sBodyArr(2) = Replace(sBodyArr(2), vbCrLf, "")
                    sBodyArr(3) = Replace(sBodyArr(3), vbCrLf, "")
                    sBodyArr(4) = Replace(sBodyArr(4), vbCrLf, "")

                    sBodyArr(1) = Trim(sBodyArr(1))
                    sBodyArr(2) = Trim(sBodyArr(2))
                    sBodyArr(3) = Trim(sBodyArr(3))
                    sBodyArr(4) = Trim(sBodyArr(4))

                    Console.WriteLine("Name: " & sBodyArr(1))
                    Console.WriteLine("Benefits Administrator: " & sBodyArr(2))
                    Console.WriteLine("Email: " & sBodyArr(3))
                    Console.WriteLine("Mobile No: " & sBodyArr(4))

                    '-------------

                    sLine01 = "Name: " & sBodyArr(1) & vbCrLf
                    sLine02 = "Benefits Administrator: " & sBodyArr(2) & vbCrLf
                    sLine03 = "Email: " & sBodyArr(3) & vbCrLf
                    sLine04 = "Mobile No: " & sBodyArr(4) & vbCrLf & vbCrLf
                    sEmailContents = sLine01 & sLine02 & sLine03 & sLine04

                    ' Create a folder named "data" under current directory
                    ' to save the email retrieved.
                    Dim localData As String = String.Format("{0}\data", Directory.GetCurrentDirectory())
                    sFileName = "Membership.txt"
                    sFullFolderPath = localData & "\" & sFileName
                    ' If the folder is not existed, create it.
                    If Not Directory.Exists(localData) Then
                        Directory.CreateDirectory(localData)
                        'Write file
                        System.IO.File.AppendAllText(sFullFolderPath, sEmailContents)
                    Else
                        'Write file
                        System.IO.File.AppendAllText(sFullFolderPath, sEmailContents)
                    End If

                Next
            End If

            Console.WriteLine("Logging out ...")
            oClient.Logout()

            Console.WriteLine("Completed!")

        Catch ep As Exception
            Console.WriteLine(ep.Message)
        End Try

    End Sub
End Module