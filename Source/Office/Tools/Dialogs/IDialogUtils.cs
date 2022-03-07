using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using NetOffice.Tools;

namespace NetOffice.OfficeApi.Tools.Dialogs
{
    /// <summary>
    /// Dialog shown event handler
    /// </summary>
    /// <param name="sender">sender instance</param>
    /// <param name="arguments">dialog shown arguments</param>
    public delegate void DialogShownEventHandler(IDialogUtils sender, DialogShownEventArgs arguments);

    /// <summary>
    /// Dialog show event handler
    /// </summary>
    /// <param name="sender">sender instance</param>
    /// <param name="arguments">dialog show arguments</param>
    public delegate void DialogShowEventHandler(IDialogUtils sender, DialogShowEventArgs arguments);

    public interface IDialogUtils
    {
        int CurrentLanguage { get; set; }
        DialogLayoutSettings Layout { get; }
        bool SuppressOnAutomation { get; set; }
        bool SuppressOnHide { get; set; }
        bool SupressGeneraly { get; set; }

        event DialogShowEventHandler DialogShow;
        event DialogShownEventHandler DialogShown;

        bool IsCurrentlySuspended();
        void ShowAbout(object modalOwner, bool modal, Size size, string headerCaption, string companyUrl, string licenceText);
        void ShowAbout(object modalOwner, bool modal, Size size, string headerCaption, string assemblyTitle, string assemblyVersion, string copyrightHint, string companyName, string companyUrl, string licenceText);
        void ShowAbout(string headerCaption, string companyUrl, string licenceText);
        void ShowDiagnostics();
        void ShowDiagnostics(bool modal);
        void ShowDiagnostics(object modalOwner, bool modal);
        void ShowDiagnostics(object modalOwner, bool modal, Size size);
        DialogResult ShowDialog(object dialog, bool modal, IEnumerable<KeyValuePair<string, object>> arguments, DialogResult defaultResult);
        DialogResult ShowDialog(object dialog, bool modal, DialogResult defaultResult);
        DialogResult ShowDialog(object modalOwner, object dialog, bool modal);
        DialogResult ShowDialog(object modalOwner, object dialogInstance, bool modal, IEnumerable<KeyValuePair<string, object>> arguments, DialogResult defaultResult);
        DialogResult ShowDialog(object modalOwner, object dialog, bool modal, DialogResult defaultResult);
        void ShowError(Exception error, string friendlyErrorDescription);
        void ShowError(object modalOwner, Exception error, string friendlyErrorDescription);
        void ShowError(object modalOwner, Exception error, string friendlyErrorDescription, bool allowDetails, bool modal, Size size);
        void ShowErrorDefault(ErrorMethodKind kind, Exception error);
        DialogResult ShowMessageBox(object modalOwner, string text, string caption, DialogButtons buttons, DialogMessageIcon icon, DialogResult defaultResult);
        DialogResult ShowMessageBox(string text, DialogMessageIcon icon, DialogResult defaultResult);
        DialogResult ShowMessageBox(string text, DialogResult defaultResult);
        DialogResult ShowMessageBox(string text, string caption, DialogButtons buttons, DialogMessageIcon icon, DialogResult defaultResult);
        DialogResult ShowMessageBox(string text, string caption, DialogButtons buttons, DialogResult defaultResult);
        DialogResult ShowMessageBox(string text, string caption, DialogMessageIcon icon, DialogResult defaultResult);
        DialogResult ShowMessageBox(string text, string caption, DialogResult defaultResult);
        DialogResult ShowText(object modalOwner, string caption, string text, string checkText, bool modal, Size size, int timeoutSeconds, bool skipOnUserAction, DialogResult defaultResult);
        DialogResult ShowText(object modalOwner, string caption, string text, string checkText, bool modal, Size size, DialogResult defaultResult);
        DialogResult ShowText(string caption, string text, int timeoutSeconds, bool skipOnUserAction, DialogResult defaultResult);
        DialogResult ShowText(string caption, string text, string checkText, bool modal, Size size, DialogResult defaultResult);
        DialogResult ShowText(string caption, string text, string checkText, DialogResult defaultResult);
    }

    /// <summary>
    /// Indicates which kind of dialog is shown
    /// </summary>
    public enum DialogType
    {
        /// <summary>
        /// Custom dialog instance
        /// </summary>
        Custom = 0,

        /// <summary>
        /// Windows.Forms MessageBox
        /// </summary>
        MessageBox = 1,

        /// <summary>
        /// Error Dialog
        /// </summary>
        Error = 2,

        /// <summary>
        /// About Dialog
        /// </summary>
        About = 3,

        /// <summary>
        /// Diagnostics Dialog
        /// </summary>
        Diagnostics = 4,

        /// <summary>
        /// Multi-Line Text Dialog, also RichText is supported
        /// </summary>
        Text = 5
    }

    /// <summary>
    /// Specifies constants defining which information to display.
    /// </summary>
    public enum DialogMessageIcon
    {
        /// <summary>
        /// The message box contain no symbols.
        /// </summary>
        None = 0,

        /// <summary>
        /// The message box contains a symbol consisting of a white X in a circle with a  red background.
        /// </summary>
        Hand = 16,

        /// <summary>
        /// The message box contains a symbol consisting of white X in a circle with a red background.
        /// </summary>
        Stop = 16,

        /// <summary>
        /// The message box contains a symbol consisting of white X in a circle with a red background.
        /// </summary>
        Error = 16,

        /// <summary>
        /// The message box contains a symbol consisting of a question mark in a circle.
        /// </summary>
        Question = 32,

        /// <summary>
        /// The message box contains a symbol consisting of an exclamation point in a triangle with a yellow background.
        /// </summary>
        Exclamation = 48,

        /// <summary>
        /// The message box contains a symbol consisting of an exclamation point in a triangle with a yellow background.
        /// </summary>
        Warning = 48,

        /// <summary>
        /// The message box contains a symbol consisting of a lowercase letter i in a circle.
        /// </summary>
        Asterisk = 64,

        /// <summary>
        /// The message box contains a symbol consisting of a lowercase letter i in a circle.
        /// </summary>
        Information = 64
    }

    /// <summary>
    /// Specifies constants defining which buttons to display
    /// </summary>
    public enum DialogButtons
    {
        /// <summary>
        /// The message box contains an OK button.
        /// </summary>
        OK = 0,

        /// <summary>
        /// The message box contains OK and Cancel buttons.
        /// </summary>
        OKCancel = 1,

        /// <summary>
        /// The message box contains Abort, Retry, and Ignore buttons.
        /// </summary>
        AbortRetryIgnore = 2,

        /// <summary>
        /// The message box contains Yes, No, and Cancel buttons.
        /// </summary>
        YesNoCancel = 3,

        /// <summary>
        /// The message box contains Yes and No buttons.
        /// </summary>
        YesNo = 4,

        /// <summary>
        /// The message box contains Retry and Cancel buttons.
        /// </summary>
        RetryCancel = 5
    }
}