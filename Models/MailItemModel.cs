namespace OutlookMailViewerMVVM.Models
{
    /// <summary>
    /// Lightweight DTO representing an Outlook email for binding.
    /// </summary>
    public sealed class MailItemModel
    {
        public System.DateTime ReceivedTime { get; init; }
        public string Subject { get; init; } = string.Empty;
        public string SenderName { get; init; } = string.Empty;
    }
}
