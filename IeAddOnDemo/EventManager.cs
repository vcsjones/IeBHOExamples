using System;
using mshtml;

namespace IeAddOnDemo
{
	public static class EventManager
	{

		public static EventProxy attachEvent(this IHTMLElement2 element, Action<CEventObj> action, string eventName)
		{
			var eventProxy = new EventProxy(eventName, element, action);
			element.attachEvent(eventName, eventProxy);
			return eventProxy;
		}

		public static EventProxy detachEvent(this IHTMLElement2 element, EventProxy eventProxy, string eventName)
		{
			element.detachEvent(eventName, eventProxy);
			return eventProxy;
		}
	}
}
