using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;

namespace OutlookMailViewerMVVM.ViewModels
{
    /// <summary>
    /// Base class with common MVVM helpers for the main screen.
    /// - INotifyPropertyChanged implementation
    /// - IsBusy & Status properties
    /// - RaiseCanExecuteChanged helper for commands
    /// - RunBusyAsync helper to wrap async actions with busy/status handling
    /// </summary>
    public abstract class MainViewModelBase : INotifyPropertyChanged
    {
        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            protected set { if (Set(ref _isBusy, value)) RaiseCanExecuteChanged(); }
        }

        private string _status = "Ready";
        public string Status
        {
            get => _status;
            protected set => Set(ref _status, value);
        }

        /// <summary>
        /// Wrap an async action with IsBusy + Status management.
        /// </summary>
        protected async Task RunBusyAsync(Func<Task> action, string? startStatus = null, string? doneStatus = null)
        {
            if (action is null) throw new ArgumentNullException(nameof(action));
            try
            {
                IsBusy = true;
                if (!string.IsNullOrWhiteSpace(startStatus)) Status = startStatus!;
                await action();
                if (!string.IsNullOrWhiteSpace(doneStatus)) Status = doneStatus!;
            }
            catch (Exception ex)
            {
                Status = "Error";
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        protected void RaiseCanExecuteChanged() => System.Windows.Input.CommandManager.InvalidateRequerySuggested();

        public event PropertyChangedEventHandler? PropertyChanged;
        protected bool Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
        {
            if (Equals(field, value)) return false;
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
            return true;
        }
    }
}
