using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows.Input;
using OutlookMailViewerMVVM.Commands;
using OutlookMailViewerMVVM.Models;
using OutlookMailViewerMVVM.Services;

namespace OutlookMailViewerMVVM.ViewModels
{
    public sealed class MainViewModel : MainViewModelBase
    {
        private readonly IOutlookService _outlookService;

        public MainViewModel() : this(new OutlookService()) { }

        public MainViewModel(IOutlookService outlookService)
        {
            _outlookService = outlookService;
            _fromDate = DateTime.Today.AddDays(-7);
            _toDate = DateTime.Today;
            LoadEmailsCommand = new AsyncRelayCommand(async _ => await LoadAsync(), _ => !IsBusy && FromDate <= ToDate);
        }

        private DateTime _fromDate;
        public DateTime FromDate
        {
            get => _fromDate;
            set { if (Set(ref _fromDate, value)) RaiseCanExecuteChanged(); }
        }

        private DateTime _toDate;
        public DateTime ToDate
        {
            get => _toDate;
            set { if (Set(ref _toDate, value)) RaiseCanExecuteChanged(); }
        }

        public ObservableCollection<MailItemModel> Emails { get; } = new();

        public ICommand LoadEmailsCommand { get; }

        private async Task LoadAsync()
        {
            await RunBusyAsync(async () =>
            {
                Emails.Clear();
                var data = await _outlookService.GetEmailsAsync(FromDate, ToDate);
                foreach (var it in data) Emails.Add(it);
            }, startStatus: "Loading from Outlook…", doneStatus: $"Loaded {Emails.Count} email(s).");
        }
    }
}
