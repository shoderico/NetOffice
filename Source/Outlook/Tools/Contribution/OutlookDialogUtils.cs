using System;

namespace NetOffice.OutlookApi.Tools.Contribution
{
    /// <summary>
    /// Outlook dialog related utils
    /// </summary>
    public class OutlookUtils
    {
        /// <summary>
        /// Try to detect the visibilty of host application main window.
        /// The implementation try to find a visible Outlook application main window and returns true if found.
        /// </summary>
        /// <param name="defaultResult">fallback result if its failed</param>
        /// <returns>true if application is visible, otherwise false</returns>
        public bool TryGetApplicationVisible(bool defaultResult)
        {
            try
            {
                Running.WindowEnumerator enumerator = new Running.WindowEnumerator("rctrl_renwnd32");
                IntPtr[] handles = enumerator.EnumerateWindows(2000);
                if (null != handles)
                {
                    // once again, no linq possible here to keep .Net2 support
                    foreach (IntPtr item in handles)
                    {
                        if (enumerator.IsVisible(item))
                            return true;
                    }
                }
                return false;
            }
            catch (System.Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
                return defaultResult;
            }
        }
    }
}