Imports System.IO
Imports System.Threading
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Calendar.v3
Imports Google.Apis.Calendar.v3.Data
Imports Google.Apis.Services
Imports Google.Apis.Util.Store

Module Program

    ' If modifying these scopes, delete your previously saved credentials
    ' at ~/.credentials/calendar-dotnet-quickstart.json
    Private ReadOnly Scopes As String() = {CalendarService.Scope.CalendarEvents}

    Private ReadOnly ApplicationName As String = "Google Calendar API .NET Quickstart"

    Sub Main()
        Dim authenticator = GetUserCredential()

        'Create Google Calendar API service
        Dim service = New CalendarService(New BaseClientService.Initializer With
        {
            .HttpClientInitializer = authenticator,
            .ApplicationName = ApplicationName
        })

        Console.WriteLine("BEFORE")
        DisplayUpcomingEvent(service)

        Console.WriteLine(Environment.NewLine)
        InsertNewEvent(service)
        Console.WriteLine(Environment.NewLine)

        Console.WriteLine("AFTER")
        DisplayUpcomingEvent(service)
    End Sub

    Private Function GetUserCredential() As UserCredential
        Using stream As New FileStream("credentials.json", FileMode.Open, FileAccess.Read)
            ' The file token.json stores the user's access and refresh tokens, and is created
            ' automatically when the authorization flow completes for the first time.
            Dim credPath As String = "token.json"

            Return GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        New FileDataStore(credPath, True)).Result
        End Using
    End Function

    Private Sub DisplayUpcomingEvent(ByVal service As CalendarService)
        Dim request As EventsResource.ListRequest = service.Events.List("primary")

        request.TimeMin = DateTime.Now
        request.ShowDeleted = False
        request.SingleEvents = True
        request.MaxResults = 20
        request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime

        ' List events.
        Dim events As Events = request.Execute()

        Console.WriteLine("Upcoming events:")

        If events.Items IsNot Nothing AndAlso events.Items.Count > 0 Then
            Dim count As Integer = 0

            For Each eventItem As [Event] In events.Items
                count += 1

                Dim [when] As String = eventItem.Start.DateTime.ToString()

                If String.IsNullOrEmpty([when]) Then
                    [when] = eventItem.Start.Date
                End If

                Console.WriteLine("{0, 2}. {1} ({2})", count, eventItem.Summary, [when])
            Next
        Else
            Console.WriteLine("No upcoming events found.")
        End If
    End Sub

    Private Sub InsertNewEvent(ByVal service As CalendarService)
        Dim newEvent As [Event] = New [Event]() With
        {
            .Summary = "Microsoft Ignite | The Tour Singapore",
            .Location = "Marina Bay Sands",
            .Description = "Learn new ways to code, optimize your cloud infrastructure, and modernize your organization with deep technical training.",
            .Start = New EventDateTime() With
            {
                .DateTime = New DateTime(2019, 1, 16, 8, 0, 0),
                .TimeZone = "Asia/Kuala_Lumpur"
            },
            .End = New EventDateTime() With
            {
                .DateTime = New DateTime(2019, 1, 17, 16, 0, 0),
                .TimeZone = "Asia/Kuala_Lumpur"
            }
        } ' optional (can be removed)
        ' .Recurrence = New String() {"RRULE:FREQ=DAILY;COUNT=2"},
        ' .Attendees = New EventAttendee() With
        ' {
        '     New EventAttendee() With {.Email = "lpage@example.com"},
        '     New EventAttendee() With {.Email = "sbrin@example.com"},
        ' },
        ' .Reminders = New [Event].RemindersData() With
        ' {
        '     .UseDefault = False,
        '     .[Overrides] = New EventReminder() With
        '     {
        '         New EventReminder() With { .Method = "email", .Minutes = 24 * 60},
        '         new EventReminder() With { .Method = "sms", .Minutes = 10 },
        '     }
        ' }
        '}
        Dim calendarId As String = "primary"
        Dim request As EventsResource.InsertRequest = service.Events.Insert(newEvent, calendarId)
        Dim createdEvent As [Event] = request.Execute()

        Console.WriteLine("Event created: {0}", createdEvent.HtmlLink)
    End Sub
End Module
