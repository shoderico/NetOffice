using System;
using System.Collections.Generic;

namespace NetOffice.OfficeApi.Tools.Dialogs
{
    /// <summary>
    /// Dialog show event arguments
    /// </summary>
    public class DialogShowEventArgs : EventArgs
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="type">dialog type</param>
        /// <param name="suppressed">dialog want not shown</param>
        /// <param name="modal">dialog want shown as modal to its parent</param>
        /// <param name="arguments">arguments dependent on dialog type</param>
        public DialogShowEventArgs(DialogType type, bool suppressed, bool modal, IEnumerable<KeyValuePair<string, object>> arguments)
        {
            Type = type;
            Suppressed = suppressed;
            Modal = modal;
            Arguments = null != arguments ? arguments : new List<KeyValuePair<string, object>>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Dialog want shown as modal to its parent.
        /// </summary>
        public bool Modal { get; private set; }

        /// <summary>
        /// The dialog want not shown because its currently forbidden by dialog settings
        /// </summary>
        public bool Suppressed { get; private set; }

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