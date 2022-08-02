Imports System.IO
Imports System.Threading
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Drive.v3
Imports Google.Apis.Util.Store

Public Class GoogleDriveService

    Private driveService As DriveService

    Public Sub New(ByVal driveService As DriveService)
        Me.driveService = driveService
    End Sub

    Public Function ListFolders() As List(Of String)
        Try
            Dim service As DriveService = driveService
            Dim listRequest As FilesResource.ListRequest = service.Files.List()
            listRequest.Fields = "nextPageToken, files(id, name)"
            listRequest.Q = $"mimeType = 'application/vnd.google-apps.folder' and trashed = false"
            Dim Folders As IList(Of Google.Apis.Drive.v3.Data.File) = listRequest.Execute().Files
            Return New List(Of String)(Folders.[Select](Function(f) "ID: " & f.Id & " - Name: " + f.Name).ToList())
        Catch __unusedException1__ As Exception
            Return Nothing
        End Try
    End Function

    Public Function ListFilesInFolder(ByVal folder As String) As List(Of String)
        Try
            Dim service As DriveService = driveService
            Dim listRequest As FilesResource.ListRequest = service.Files.List()
            listRequest.Q = $"mimeType != 'application/vnd.google-apps.folder' and trashed = false and '{Find_IdFile_ForName(folder)(0)}' in parents"
            Dim files As IList(Of Google.Apis.Drive.v3.Data.File) = listRequest.Execute().Files
            Return New List(Of String)(files.[Select](Function(f) "ID: " & f.Id & " - Name: " + f.Name).ToList())
        Catch __unusedException1__ As Exception
            Return Nothing
        End Try
    End Function

    Public Function ListAllFiles() As List(Of String)
        Try
            Dim service As DriveService = driveService
            Dim listRequest As FilesResource.ListRequest = service.Files.List()
            listRequest.Q = $"mimeType != 'application/vnd.google-apps.folder' and trashed = false"
            Dim files As IList(Of Google.Apis.Drive.v3.Data.File) = listRequest.Execute().Files
            Return New List(Of String)(files.[Select](Function(f) "ID: " & f.Id & " - Name: " + f.Name).ToList())
        Catch __unusedException1__ As Exception
            Return Nothing
        End Try
    End Function

    Public Function ListTrashFiles() As List(Of String)
        Try
            Dim service As DriveService = driveService
            Dim listRequest As FilesResource.ListRequest = service.Files.List()
            listRequest.Q = $"mimeType != 'application/vnd.google-apps.folder' and trashed = true"
            Dim files As IList(Of Google.Apis.Drive.v3.Data.File) = listRequest.Execute().Files
            Return New List(Of String)(files.[Select](Function(f) "ID: " & f.Id & " - Name: " + f.Name).ToList())
        Catch __unusedException1__ As Exception
            Return Nothing
        End Try
    End Function

    Public Function UploadFile(ByVal descripcion As String, ByVal filepath As String, ByVal Optional folder As String = "My Drive") As String
        Try
            Dim service As DriveService = driveService
            Dim request As FilesResource.CreateMediaUpload
            Dim driveFile = New Google.Apis.Drive.v3.Data.File() With {
                .Name = Path.GetFileName(filepath),
                .Description = descripcion,
                .MimeType = GetMimeType(filepath),
                .Parents = Find_IdFile_ForName(folder)
            }

            Using stream = System.IO.File.OpenRead(filepath)
                request = service.Files.Create(driveFile, stream, driveFile.MimeType)
                request.Fields = "id"
                Dim response = request.Upload()

                Try
                    Dim file = request.ResponseBody
                    Return file.Id
                Catch __unusedException1__ As Exception
                    Return "The name folder not exist"
                End Try
            End Using

        Catch __unusedException1__ As Exception
            Return Nothing
        End Try
    End Function

    Public Function CreateFolder(ByVal name As String) As String
        Try
            Dim service As DriveService = driveService
            Dim driveFolder = New Google.Apis.Drive.v3.Data.File() With {
                .Name = name,
                .MimeType = "application/vnd.google-apps.folder"
            }
            Dim request = service.Files.Create(driveFolder)
            Dim response = request.Execute()
            Return response.Id
        Catch __unusedException1__ As Exception
            Return Nothing
        End Try
    End Function

    Private Function GetMimeType(ByVal FileName As String) As String

        Dim mimeType As String = "application/unknown"
        Dim ext As String = System.IO.Path.GetExtension(FileName).ToLower()
        Dim regKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext)
        If regKey IsNot Nothing AndAlso regKey.GetValue("Content Type") IsNot Nothing Then
            mimeType = regKey.GetValue("Content Type").ToString()
        End If
        System.Diagnostics.Debug.WriteLine(mimeType)
        Return mimeType
    End Function

    Private Function Find_IdFile_ForName(ByVal name As String) As List(Of String)
        Try
            Dim service As DriveService = driveService
            Dim request As FilesResource.ListRequest = service.Files.List()
            request.Q = $"name = '{name}'"
            Dim files As IList(Of Google.Apis.Drive.v3.Data.File) = request.Execute().Files
            Return New List(Of String)(files.[Select](Function(f) f.Id).ToList())
        Catch __unusedException1__ As Exception
            Return Nothing
        End Try
    End Function
End Class

Public Class Auth
    Private pathClientSecretJson As String = ""
    Shared Scopes As String() = {DriveService.Scope.Drive, DriveService.Scope.DriveFile}
    Shared ApplicationName As String = ""
    Private credential As UserCredential

    Public Sub New(ByVal pathClientSecretJson As String, ByVal pathForSaveCredentials As String, ByVal appName As String)
        Me.pathClientSecretJson = pathClientSecretJson
        ApplicationName = appName
        SaveLocalCredentials(pathForSaveCredentials)
    End Sub

    Private Sub SaveLocalCredentials(ByVal pathCredentials As String)
        Using stream = New FileStream(pathClientSecretJson, FileMode.Open, FileAccess.Read)
            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.FromStream(stream).Secrets, Scopes, "user", CancellationToken.None, New FileDataStore(pathCredentials, True)).Result
            Console.WriteLine("Creadenciales guardadas en: " & pathCredentials)
        End Using
    End Sub

    Public Function MakeService() As DriveService
        Dim service = New DriveService(New Google.Apis.Services.BaseClientService.Initializer With {
            .HttpClientInitializer = credential,
            .ApplicationName = ApplicationName
        })
        Return service
    End Function
End Class


