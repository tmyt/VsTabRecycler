//------------------------------------------------------------------------------
// <copyright file="VsTabRecyclerPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;

namespace VsTabRecycler
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string)]
    public sealed class VsTabRecyclerPackage : Package
    {
        /// <summary>
        /// VsTabRecyclerPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "27474c93-3eda-4cae-80dd-2363808c4592";

        private const int LRUMaxNum = 10;

        private List<Window> _lruList;

        /// <summary>
        /// Initializes a new instance of the <see cref="VsTabRecyclerPackage"/> class.
        /// </summary>
        public VsTabRecyclerPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            DTE dte = (DTE)this.GetService(typeof(DTE));
            dte.Events.WindowEvents.WindowActivated += WindowEvents_WindowActivated;
            dte.Events.WindowEvents.WindowCreated += WindowEvents_WindowCreated;
            dte.Events.WindowEvents.WindowClosing += WindowEvents_WindowClosing;
        }

        private void WindowEvents_WindowClosing(Window Window)
        {
            if (Window?.Kind != "Document") return;
            //Debug.WriteLine("Window Closing: " + Window.Document?.FullName);
            _lruList.Remove(Window);
        }

        private void WindowEvents_WindowCreated(Window Window)
        {
            if (Window?.Kind != "Document") return;
            //Debug.WriteLine("Window Created: " + Window.Document?.FullName);
            ScanWindows();
            while (_lruList.Count > LRUMaxNum)
            {
                var n = _lruList.Count;
                try
                {
                    _lruList[0].Close();
                }
                catch
                { }
                // 一番古いドキュメントが変更済みでなんか閉じれなかったから閉じるのやめる
                if (n == _lruList.Count)
                {
                    break;
                }
            }
        }

        private void WindowEvents_WindowActivated(Window GotFocus, Window LostFocus)
        {
            if (GotFocus?.Kind != "Document") return;
            //Debug.WriteLine("Window Activated: " + GotFocus.Document?.FullName);
            if (_lruList.All(w => w != GotFocus))
            {
                _lruList.Add(GotFocus);
            }
            _lruList.Remove(GotFocus);
            _lruList.Add(GotFocus);
        }

        private void ScanWindows()
        {
            var dte = (DTE)this.GetService(typeof(DTE));
            foreach (var d in dte.Windows.OfType<Window>())
            {
                if ((d).Kind != "Document") continue;
                if (_lruList.All(w => w != d))
                {
                    _lruList.Add(d);
                }
            }
        }

        #endregion
    }
}
