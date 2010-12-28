using System.Runtime.InteropServices;
using mshtml;

namespace IeAddOnDemo
{
	[ComVisible(true),
	Guid("9AB12757-BDAF-4F9A-8DE8-413C3615590C"),
	ClassInterface(ClassInterfaceType.None)]
	public partial class BHO : IObjectWithSite
	{
		partial void SetSiteStartup()
		{
            _webBrowser2Events.DocumentComplete += _webBrowser2Events_DocumentComplete;

		}
		partial void SetSiteShutdown()
		{
            _webBrowser2Events.DocumentComplete -= _webBrowser2Events_DocumentComplete;

		}

        private void _webBrowser2Events_DocumentComplete(object pdisp, ref object url)
        {
            if (!ReferenceEquals(pdisp, _pUnkSite))
            {
                return;
            }
            var document = (HTMLDocument) _pUnkSite.Document;
            var imageElements = document.getElementsByTagName("img");
            foreach (IHTMLElement htmlImgElement in imageElements)
            {
                htmlImgElement.style.borderColor = "green";
                htmlImgElement.style.borderWidth = "3px";
                htmlImgElement.style.borderStyle = "solid";
            }
        }
	}
}
