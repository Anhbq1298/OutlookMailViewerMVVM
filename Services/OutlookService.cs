using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookMailViewerMVVM.Models;

namespace OutlookMailViewerMVVM.Services
{
    /// <summary>
    /// Outlook interop wrapper. Reads from default Inbox.
    /// Make sure Outlook (desktop) is installed and a profile is configured.
    /// </summary>
    public sealed class OutlookService : IOutlookService
    {
        public Task<IReadOnlyList<MailItemModel>> GetEmailsAsync(DateTime from, DateTime to)
        {
            return Task.Run(() =>
            {
                var results = new List<MailItemModel>();
                Outlook.Application? app = null;
                Outlook.NameSpace? ns = null;
                Outlook.MAPIFolder? inbox = null;
                Outlook.Items? items = null;
                Outlook.Items? restricted = null;

                // Outlook restrict filter requires US culture date format
                var us = new CultureInfo("en-US");
                var fromStr = from.ToString("g", us);
                // Add 23:59:59 to include the entire 'to' day if the user selected date-only
                var toInclusive = to.Date.AddDays(1).AddSeconds(-1);
                var toStr = toInclusive.ToString("g", us);

                string filter = $"[ReceivedTime] >= '{fromStr}' AND [ReceivedTime] <= '{toStr}'";

                try
                {
                    app = new Outlook.Application();
                    ns = app.GetNamespace("MAPI");
                    inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    items = inbox.Items;
                    items.Sort("[ReceivedTime]", true); // newest first
                    items.IncludeRecurrences = true;

                    restricted = items.Restrict(filter);

                    foreach (var obj in restricted)
                    {
                        if (obj is Outlook.MailItem mail)
                        {
                            string senderName = string.Empty;
                            try
                            {
                                senderName = mail.SenderName ?? string.Empty;
                            }
                            catch { /* ignore */ }

                            results.Add(new MailItemModel
                            {
                                ReceivedTime = mail.ReceivedTime,
                                Subject = mail.Subject ?? string.Empty,
                                SenderName = senderName
                            });

                            Marshal.FinalReleaseComObject(mail);
                        }
                    }
                }
                finally
                {
                    if (restricted != null) Marshal.FinalReleaseComObject(restricted);
                    if (items != null) Marshal.FinalReleaseComObject(items);
                    if (inbox != null) Marshal.FinalReleaseComObject(inbox);
                    if (ns != null) Marshal.FinalReleaseComObject(ns);
                    if (app != null) Marshal.FinalReleaseComObject(app);
                }

                return (IReadOnlyList<MailItemModel>)results;
            });
        }
    }
}
