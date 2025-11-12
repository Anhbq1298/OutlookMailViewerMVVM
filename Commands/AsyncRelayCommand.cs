using System;
using System.Threading.Tasks;
using System.Windows.Input;

namespace OutlookMailViewerMVVM.Commands
{
    /// <summary>
    /// Async command wrapper to avoid async-void in MVVM bindings.
    /// Disables execution while a previous task is running.
    /// </summary>
    public sealed class AsyncRelayCommand : ICommand
    {
        private readonly Func<object?, Task> _executeAsync;
        private readonly Func<object?, bool>? _canExecute;
        private bool _isRunning;

        public AsyncRelayCommand(Func<object?, Task> executeAsync, Func<object?, bool>? canExecute = null)
        {
            _executeAsync = executeAsync ?? throw new ArgumentNullException(nameof(executeAsync));
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object? parameter)
        {
            if (_isRunning) return false;
            return _canExecute?.Invoke(parameter) ?? true;
        }

        public async void Execute(object? parameter)
        {
            if (!CanExecute(parameter)) return;
            try
            {
                _isRunning = true;
                CommandManager.InvalidateRequerySuggested();
                await _executeAsync(parameter);
            }
            finally
            {
                _isRunning = false;
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }
}
