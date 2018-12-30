using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.Threading;

namespace Zephyr.TestDrive.GoogleCalendarApi
{
    internal class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        private static readonly string[] Scopes = { CalendarService.Scope.CalendarEvents };

        private static readonly string ApplicationName = "Google Calendar API .NET Quickstart";

        private static void Main(string[] args)
        {
            var authenticator = GetUserCredential();

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = authenticator,
                ApplicationName       = ApplicationName,
            });

            Console.WriteLine("BEFORE");
            DisplayUpcomingEvent(service);

            Console.WriteLine(Environment.NewLine);
            InsertNewEvent(service);
            Console.WriteLine(Environment.NewLine);

            Console.WriteLine("AFTER");
            DisplayUpcomingEvent(service);
        }

        private static UserCredential GetUserCredential()
        {
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                var credPath = "token.json";

                return GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
            }
        }

        private static void DisplayUpcomingEvent(CalendarService service)
        {
            EventsResource.ListRequest request = service.Events.List("primary");

            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 20;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // List events.
            var events = request.Execute();

            Console.WriteLine("Upcoming events:");

            if (events.Items != null && events.Items.Count > 0)
            {
                int count = 0;

                foreach (var eventItem in events.Items)
                {
                    count++;

                    string when = eventItem.Start.DateTime.ToString();

                    if (String.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                    }

                    Console.WriteLine("{0, 2}. {1} ({2})", count, eventItem.Summary, when);
                }
            }
            else
            {
                Console.WriteLine("No upcoming events found.");
            }
        }

        private static void InsertNewEvent(CalendarService service)
        {
            var newEvent = new Event()
            {
                Summary     = "Microsoft Ignite | The Tour Singapore",
                Location    = "Marina Bay Sands",
                Description = "Learn new ways to code, optimize your cloud infrastructure, and modernize your organization with deep technical training.",
                Start = new EventDateTime()
                {
                    DateTime = new DateTime(2019, 1, 16, 8, 0, 0),
                    TimeZone = "Asia/Kuala_Lumpur"
                },
                End = new EventDateTime()
                {
                    DateTime = new DateTime(2019, 1, 17, 16, 0, 0),
                    TimeZone = "Asia/Kuala_Lumpur"
                },
                //Recurrence = new String[] { "RRULE:FREQ=DAILY;COUNT=2" },
                //Attendees = new EventAttendee[]
                //{
                //    new EventAttendee() { Email = "lpage@example.com" },
                //    new EventAttendee() { Email = "sbrin@example.com" },
                //},
                //Reminders = new Event.RemindersData()
                //{
                //    UseDefault = false,
                //    Overrides = new EventReminder[]
                //    {
                //        new EventReminder() { Method = "email", Minutes = 24 * 60 },
                //        new EventReminder() { Method = "sms", Minutes = 10 },
                //    }
                //}
            };

            var calendarId   = "primary";
            var request      = service.Events.Insert(newEvent, calendarId);
            var createdEvent = request.Execute();

            Console.WriteLine("Event created: {0}", createdEvent.HtmlLink);
        }
    }
}
