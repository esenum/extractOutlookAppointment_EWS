using Microsoft.Exchange.WebServices.Data;
using System;
using System.Net;
using System.Security.Policy;

namespace AccessToCalendarEWS
{
    internal class Program
    {

        public static ExchangeService ConnectToService(string UserName, string UserPass, ExchangeVersion exchangeVersion, string Url, ITraceListener listener)
        {
            ExchangeService service = new ExchangeService(exchangeVersion);

            if (listener != null)
            {
                service.TraceListener = listener;
                service.TraceFlags = TraceFlags.All;
                service.TraceEnabled = true;
            }

            service.Credentials = new NetworkCredential(UserName, UserPass);

            Console.Write(string.Format("Connect EWS URL for {0}. Please wait... ", UserName));

            service.Url = new Uri(Url);

            Console.WriteLine("Complete");
            return service;
        }
        static void Main(string[] args)
        {
            ExchangeService service = ConnectToService("yourMailAddress", "Password", ExchangeVersion.Exchange2013, "https://outlook.office365.com/EWS/Exchange.asmx", null);

            ServicePointManager.ServerCertificateValidationCallback = (sender, cert, chain, sslPolicyErrors) => true;

            EmailAddressCollection myRoomLists = service.GetRoomLists();
            foreach (EmailAddress myAddress in myRoomLists)
            {
                Console.WriteLine("Room name: {0}", myAddress.Name);
                System.Collections.ObjectModel.Collection<EmailAddress> myRoomAddresses = service.GetRooms(myAddress);

                foreach (EmailAddress address in myRoomAddresses)
                {
                    Console.WriteLine("Email Address: {0}", address.Address);
                    Console.WriteLine("Name: {0}", address.Name);
                }
            }

           

            try
            {
                String MailboxToAccess = "";             

                DateTime startDate = new DateTime(2022,9, 20);
                DateTime endDate = new DateTime(2022, 9, 30);
                CalendarView calView = new CalendarView(startDate, endDate);             
                
                FolderId CalendarFolderId = new FolderId(WellKnownFolderName.Calendar, MailboxToAccess);

                FindItemsResults<Item> instanceResults = service.FindItems(CalendarFolderId, calView);
                service.LoadPropertiesForItems(instanceResults, new PropertySet(
                 ItemSchema.Subject,
                 AppointmentSchema.Start,
                 AppointmentSchema.End,
                // AppointmentSchema.IsAllDayEvent,
                 AppointmentSchema.Organizer,
                 AppointmentSchema.Location,
                 AppointmentSchema.RequiredAttendees,
                 AppointmentSchema.OptionalAttendees,
                 //ItemSchema.TextBody,
                 //ItemSchema.ReminderMinutesBeforeStart,
                 //ItemSchema.DisplayTo,
                 //ItemSchema.DisplayCc,
                // AppointmentSchema.IsRecurring,
                  AppointmentSchema.IsCancelled,
                // AppointmentSchema.IsOnlineMeeting,
                 AppointmentSchema.AppointmentType,
                 ItemSchema.Sensitivity));



                foreach (Item item in instanceResults.Items)
                {
                    Appointment appointment = item as Appointment;

                    
                        Console.WriteLine("------------------------------------------------------------------");

                        //Console.WriteLine("IsCanceled : " + appointment.IsCancelled);
                        Console.WriteLine("Subject : " + appointment.Subject);

                        if (appointment.IsCancelled)
                        {
                            Console.WriteLine("Status : Canceled");
                        }
                        else
                        {
                            Console.WriteLine("Status : Active");
                        }

                        Console.WriteLine("Organizer : " + appointment.Organizer);
                        Console.WriteLine("Organizer Email : " + appointment.Organizer.Address);
                        Console.WriteLine("Start : " + appointment.Start);
                        Console.WriteLine("End : " + appointment.End);

                        TimeSpan ts = appointment.End - appointment.Start;
                        double ToplamGun = Convert.ToDouble(ts.Days);
                        double ToplamSaat = Convert.ToDouble(ts.Hours);
                        double ToplamDakika = Convert.ToDouble(ts.Minutes);

                        

                        if (ts.Days > 0 && ts.Hours > 0 && ts.Minutes > 0)
                            Console.WriteLine("Duration : {0} gün {1} saat {2} dakika ", ts.Days, ts.Hours, ts.Minutes);
                        else if (ts.Days > 0 && ts.Hours == 0 && ts.Minutes > 0)
                            Console.WriteLine("Duration : {0} gün  {1} dakika ", ts.Days, ts.Minutes);
                        else if (ts.Days > 0 && ts.Hours > 0 && ts.Minutes == 0)
                            Console.WriteLine("Duration : {0} gün {1} saat ", ts.Days, ts.Hours);
                        else if (ts.Days > 0 && ts.Hours == 0 && ts.Minutes == 0)
                            Console.WriteLine("Duration : {0} gün ", ts.Days);
                        else if (ts.Days == 0 && ts.Hours > 0 && ts.Minutes > 0)
                            Console.WriteLine("Duration : {0} saat {1} dakika ", ts.Hours, ts.Minutes);
                        else if (ts.Days == 0 && ts.Hours > 0 && ts.Minutes == 0)
                            Console.WriteLine("Duration : {0} saat  ", ts.Hours);
                        else if (ts.Days == 0 && ts.Hours == 0 && ts.Minutes > 0)
                            Console.WriteLine("Duration : {0} dakika ", ts.Minutes);

                        Console.WriteLine("Location : " + appointment.Location);

                        //Console.WriteLine("AllDay : " + appointment.IsAllDayEvent);
                        //Console.WriteLine("Body : " + appointment.TextBody);
                        //Console.WriteLine("Reminder : " + appointment.ReminderMinutesBeforeStart);
                        //Console.WriteLine("DisplayTo : " + appointment.DisplayTo);
                        //Console.WriteLine("DisplayCc : " + appointment.DisplayCc);



                        //Console.WriteLine("Reminder : " + appointment.ReminderMinutesBeforeStart);

                        string katilimcilar = "";
                        for (int j = 0; j < appointment.RequiredAttendees.Count; j++)
                        {
                            katilimcilar = katilimcilar + appointment.RequiredAttendees[j].Address + ";";

                        }
                        katilimcilar = "Required Attendee Email :" + katilimcilar;
                        Console.Write(katilimcilar);

                        //string katilimcilar = "";
                        //foreach (var word in appointment.RequiredAttendees[].Address.Split(';'))
                        //{
                        //    Console.WriteLine("Required Attendee Email: {0}", word);
                        //}



                        for (int j = 0; j < appointment.OptionalAttendees.Count; j++)
                        {
                            Console.WriteLine("Optional Attendee Email : " + appointment.OptionalAttendees[j].Address);

                        }

                        if (appointment.AppointmentType == AppointmentType.Exception)
                        {
                            Console.WriteLine("Toplantı belirli periyotlarla tekrar edilmektedir.");
                        }
                        else
                        {
                            Console.WriteLine("Toplantı tekrar eden tipte değildir.");
                        }
                    

                    
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }



    }
}