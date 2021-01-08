// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using GraphTutorial.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Web;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TimeZoneConverter;

namespace GraphTutorial.Controllers
{
    public class CalendarController : Controller
    {
        private readonly GraphServiceClient _graphClient;
        private readonly ILogger<HomeController> _logger;

        public CalendarController(
            GraphServiceClient graphClient,
            ILogger<HomeController> logger)
        {
            _graphClient = graphClient;
            _logger = logger;
        }

        // <IndexSnippet>
        // Minimum permission scope needed for this view
        [AuthorizeForScopes(Scopes = new[] { "Calendars.Read" })]
        public async Task<IActionResult> Index()
        {
            try
            {
                var userTimeZone = TZConvert.GetTimeZoneInfo(
                    User.GetUserGraphTimeZone());
                var startOfWeek = CalendarController.GetUtcStartOfWeekInTimeZone(
                    DateTime.Today, userTimeZone);

                var events = await GetUserWeekCalendar(startOfWeek);

                var model = new CalendarViewModel(startOfWeek, events);

                return View(model);
            }
            catch (ServiceException ex)
            {
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException)
                {
                    throw ex;
                }

                return View(new CalendarViewModel())
                    .WithError("Error getting calendar view", ex.Message);
            }
        }
        // </IndexSnippet>

        // <CalendarNewGetSnippet>
        // Minimum permission scope needed for this view
        [AuthorizeForScopes(Scopes = new[] { "Calendars.ReadWrite" })]
        public IActionResult New()
        {
            return View();
        }
        // </CalendarNewGetSnippet>

        // <CalendarNewPostSnippet>
        [HttpPost]
        [ValidateAntiForgeryToken]
        [AuthorizeForScopes(Scopes = new[] { "Calendars.ReadWrite" })]
        public async Task<IActionResult> New([Bind("Subject,Attendees,Start,End,Body")] NewEvent newEvent)
        {

            var timeZone = User.GetUserGraphTimeZone();
            //Getting the free/busy shedule time
            var schedules = new List<String>()
               {
                "kondaseetharam@hotmail.com",
                "sk00519964@techmahindra.com"
                };
            var Start = new DateTimeTimeZone
            {
                DateTime = newEvent.Start.ToString("o"),
                // Use the user's time zone
                TimeZone = timeZone
            };
            var End = new DateTimeTimeZone
            {
                DateTime = newEvent.End.ToString("o"),
                // Use the user's time zone
                TimeZone = timeZone
            };
         
            var availabilityViewInterval = 60;
            var data = await _graphClient.Me.Calendar
                                            .GetSchedule(schedules, End, Start, availabilityViewInterval)
                                            .Request()
                                            .Header("Prefer", $"outlook.timezone=\"{User.GetUserGraphTimeZone()}\"")
                                            .PostAsync();
            string meetingStartTime = string.Empty;

            IEnumerable<ScheduleItem> items = data.CurrentPage[0].ScheduleItems;

            foreach (var item in items)
            {
                meetingStartTime = item.End.DateTime.ToString();
            }
            if (String.IsNullOrEmpty(meetingStartTime))
            {
                // string todayDate = // DateTime.Now.ToString("yyyy-MM-dd") + "T12:00:00.0000000";
                meetingStartTime = newEvent.Start.ToString("o");
            }
            else
            {
                var date1 = DateTime.Parse(meetingStartTime);
                var date2 = DateTime.Parse(newEvent.End.ToString("o"));
                if (date1 >= date2)
                {
                    return RedirectToAction("Index").WithSuccess("There is no free slots b/w 12pm and 4pm");
                }
            }

            DateTime objDateTime = DateTime.Parse(meetingStartTime);
            //Adding 1 hour to meeting Start Time
            string meetingEndTime = objDateTime.AddHours(1).ToString("yyyy-MM-dd'T'HH:mm:ss.0000000");
            // Create a Graph event with the required fields
            var graphEvent = new Event
            {
                Subject = newEvent.Subject,
                Start = new DateTimeTimeZone
                {
                    DateTime = meetingStartTime,// newEvent.Start.ToString("o"),
                                                // Use the user's time zone
                    TimeZone = timeZone
                },
                End = new DateTimeTimeZone
                {
                    DateTime = meetingEndTime,//newEvent.End.ToString("o"),
                                              // Use the user's time zone
                    TimeZone = timeZone
                }
            };

            // Add body if present
            if (!string.IsNullOrEmpty(newEvent.Body))
            {
                graphEvent.Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = newEvent.Body
                };
            }

            // Add attendees if present
            if (!string.IsNullOrEmpty(newEvent.Attendees))
            {
                var attendees =
                    newEvent.Attendees.Split(';', StringSplitOptions.RemoveEmptyEntries);

                if (attendees.Length > 0)
                {
                    var attendeeList = new List<Attendee>();
                    foreach (var attendee in attendees)
                    {
                        attendeeList.Add(new Attendee
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = attendee
                            },
                            Type = AttendeeType.Required
                        });
                    }
                }
            }

            try
            {


                //Add the event
                await _graphClient.Me.Events
                    .Request()
                     .AddAsync(graphEvent);


                // Redirect to the calendar view with a success message
                return RedirectToAction("Index").WithSuccess("Event created");
            }
            catch (ServiceException ex)
            {
                // Redirect to the calendar view with an error message
                return RedirectToAction("Index")
                    .WithError("Error creating event", ex.Error.Message);
            }

        }
        // </CalendarNewPostSnippet>

        // <GetCalendarViewSnippet>
        private async Task<IList<Event>> GetUserWeekCalendar(DateTime startOfWeek)
        {
            // Configure a calendar view for the current week
            var endOfWeek = startOfWeek.AddDays(7);

            var viewOptions = new List<QueryOption>
            {
                new QueryOption("startDateTime", startOfWeek.ToString("o")),
                new QueryOption("endDateTime", endOfWeek.ToString("o"))
            };

            var events = await _graphClient.Me
                .CalendarView
                .Request(viewOptions)
                // Send user time zone in request so date/time in
                // response will be in preferred time zone
                .Header("Prefer", $"outlook.timezone=\"{User.GetUserGraphTimeZone()}\"")
                // Get max 50 per request
                .Top(50)
                // Only return fields app will use
                .Select(e => new
                {
                    e.Subject,
                    e.Organizer,
                    e.Start,
                    e.End
                })
                // Order results chronologically
                .OrderBy("start/dateTime")
                .GetAsync();

            IList<Event> allEvents;
            // Handle case where there are more than 50
            if (events.NextPageRequest != null)
            {
                allEvents = new List<Event>();
                // Create a page iterator to iterate over subsequent pages
                // of results. Build a list from the results
                var pageIterator = PageIterator<Event>.CreatePageIterator(
                    _graphClient, events,
                    (e) =>
                    {
                        allEvents.Add(e);
                        return true;
                    }
                );
                await pageIterator.IterateAsync();
            }
            else
            {
                // If only one page, just use the result
                allEvents = events.CurrentPage;
            }

            return allEvents;
        }

        private static DateTime GetUtcStartOfWeekInTimeZone(DateTime today, TimeZoneInfo timeZone)
        {
            // Assumes Sunday as first day of week
            int diff = System.DayOfWeek.Sunday - today.DayOfWeek;

            // create date as unspecified kind
            var unspecifiedStart = DateTime.SpecifyKind(today.AddDays(diff), DateTimeKind.Unspecified);

            // convert to UTC
            return TimeZoneInfo.ConvertTimeToUtc(unspecifiedStart, timeZone);
        }
        private async Task<MeetingTimeSuggestionsResult> FindMeetingSuggestions(DateTimeTimeZone start,DateTimeTimeZone end)
        {

            
            var timeZone = User.GetUserGraphTimeZone();

            var attendees = new List<AttendeeBase>()
                            {
                                new AttendeeBase
                                {
                                    Type = AttendeeType.Required,
                                    EmailAddress = new EmailAddress
                                    {
                                        Name = "seetharam konda",
                                        Address = "kondaseetharam@hotmail.com"
                                    }
                                }
                            };

            var locationConstraint = new LocationConstraint
                                    {
                                        IsRequired = false,
                                        SuggestLocation = false,
                                        Locations = new List<LocationConstraintItem>()
                                    {
                                        new LocationConstraintItem
                                        {
                                            ResolveAvailability = false,
                                            DisplayName = "Conf room Hood"
                                        }
                                     }
                                   };

            var timeConstraint = new TimeConstraint
            {
                ActivityDomain = ActivityDomain.Work,
                TimeSlots = new List<TimeSlot>()
                             {
                                new TimeSlot
                                {
                                    Start=start,
                                    End=end
                                    //Start = new DateTimeTimeZone
                                    //{
                                    //    DateTime = "2019-04-16T09:00:00",
                                    //    TimeZone = "Pacific Standard Time"
                                    //},
                                    //End = new DateTimeTimeZone
                                    //{
                                    //    DateTime = "2019-04-18T17:00:00",
                                    //    TimeZone = "Pacific Standard Time"
                                    //}
                                }
                            }
            };

            var isOrganizerOptional = false;

            var meetingDuration = new Duration("PT1H");

            var returnSuggestionReasons = true;

            var minimumAttendeePercentage = (double)100;

            MeetingTimeSuggestionsResult x = await _graphClient.Me
                  .FindMeetingTimes(attendees, locationConstraint, timeConstraint, meetingDuration, null, isOrganizerOptional, returnSuggestionReasons, minimumAttendeePercentage)
                  .Request()
                  .Header("Prefer", "outlook.timezone=\"India Standard Time\"")
                  .PostAsync();
            return x;
        }
        // </GetCalendarViewSnippet>
    }
}
