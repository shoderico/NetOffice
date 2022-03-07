using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace NetOffice.OfficeApi.Tools.Dialogs
{
    /// <summary>
    /// Dialog shown event arguments
    /// </summary>
    public class DialogShownEventArgs : EventArgs
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="type">dialog type</param>
        /// <param name="suppressed">dialog has not shown</param>
        /// <param name="modal">dialog has shown as modal to its parent</param>
        /// <param name="result">dialog result if set</param>
        /// <param name="arguments">arguments dependent on dialog type</param>
        public DialogShownEventArgs(DialogType type, bool suppressed, bool modal, DialogResult result, IEnumerable<KeyValuePair<string, object>> arguments)
        {
            Type = type;
            Suppressed = suppressed;
            Modal = modal;
            Result = result;
            Arguments = null != arguments ? arguments : new List<KeyValuePair<string, object>>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Dialog has shown as modal to its parent.
        /// </summary>
        public bool Modal { get; private set; }

        /// <summary>
        /// The dialog has not shown because its currently forbidden by dialog settings
        /// </summary>
        public bool Suppressed { get; private set; }

        /// <summary>
        /// Dialog result if set
        /// </summary>
        public DialogResult Result { get; private set; }

        /// <summary>
        /// Shown dialog type
        /// </summary>
        public DialogType Type { get; private set; }

        /// <summary>
        /// Arguments dependent on dialog type
        /// </summary>
        public IEnumerable<KeyValuePair<string, object>> Arguments { get; private set; }

        #endregion
    }
}