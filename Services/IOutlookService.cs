using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OutlookMailViewerMVVM.Models;

namespace OutlookMailViewerMVVM.Services
{
    public interface IOutlookService
    {
        /// <summary>
        /// Returns emails in the Inbox received between [from, to].
        /// NOTE: Requires Outlook installed and a configured profile.
        /// </summary>
        Task<IReadOnlyList<MailItemModel>> GetEmailsAsync(DateTime from, DateTime to);
    }
}
