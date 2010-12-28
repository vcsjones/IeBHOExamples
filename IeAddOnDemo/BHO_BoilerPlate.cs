using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using SHDocVw;

namespace IeAddOnDemo
{
    partial class BHO
    {
        private IWebBrowser2 _pUnkSite;
        private DWebBrowserEvents2_Event _webBrowser2Events;
        private readonly List<EventProxy> _wiredEvents = new List<EventProxy>();

        public int SetSite(object pUnkSite)
        {
            if (pUnkSite != null)
            {
                _pUnkSite = (IWebBrowser2) pUnkSite;
                _webBrowser2Events = (DWebBrowserEvents2_Event) pUnkSite;
                SetSiteStartup();
                _webBrowser2Events.BeforeNavigate2 += _webBrowser2Events_BeforeNavigate2;
            }
            else
            {
                SetSiteShutdown();
                _webBrowser2Events.BeforeNavigate2 -= _webBrowser2Events_BeforeNavigate2;
                _pUnkSite = null;
                _webBrowser2Events = null;
            }
            return 0;
        }

        void _webBrowser2Events_BeforeNavigate2(object pDisp, ref object URL, ref object Flags,
        ref object TargetFrameName, ref object PostData, ref object Headers, ref bool Cancel)
        {
            foreach (var wiredEvent in _wiredEvents)
            {
                wiredEvent.Target.detachEvent(wiredEvent, wiredEvent.EventName);
            }
            _wiredEvents.Clear();
        }

        public int GetSite(ref Guid riid, out IntPtr ppvSite)
        {
            var pUnk = Marshal.GetIUnknownForObject(_pUnkSite);
            try
            {
                return Marshal.QueryInterface(pUnk, ref riid, out ppvSite);
            }
            finally
            {
                Marshal.Release(pUnk);
            }
        }

        partial void SetSiteStartup();
        partial void SetSiteShutdown();
    }
}
