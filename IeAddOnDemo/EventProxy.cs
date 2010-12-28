using System;
using System.Globalization;
using System.Reflection;
using mshtml;

namespace IeAddOnDemo
{
	public class EventProxy : IReflect, IEquatable<EventProxy>
	{
		private readonly string _eventName;
		private readonly IHTMLElement2 _target;
		private readonly Action<CEventObj> _eventHandler;
		private readonly Type _type;

		public EventProxy(string eventName, IHTMLElement2 target, Action<CEventObj> eventHandler)
		{
			_eventName = eventName;
			_target = target;
			_eventHandler = eventHandler;
			_type = typeof(EventProxy);
		}

		public IHTMLElement2 Target
		{
			get { return _target; }
		}

		public Action<CEventObj> Handler
		{
			get { return _eventHandler; }
		}

		public string EventName
		{
			get { return _eventName; }
		}

		public void OnHtmlEvent(object o)
		{
			InvokeClrEvent((CEventObj)o);
		}

		private void InvokeClrEvent(CEventObj o)
		{
			if (_eventHandler != null)
			{
				_eventHandler(o);
			}
		}

		public MethodInfo GetMethod(string name, BindingFlags bindingAttr, Binder binder, Type[] types, ParameterModifier[] modifiers)
		{
			return _type.GetMethod(name, bindingAttr, binder, types, modifiers);
		}

		public MethodInfo GetMethod(string name, BindingFlags bindingAttr)
		{
			return _type.GetMethod(name, bindingAttr);
		}

		public MethodInfo[] GetMethods(BindingFlags bindingAttr)
		{
			return _type.GetMethods(bindingAttr);
		}

		public FieldInfo GetField(string name, BindingFlags bindingAttr)
		{
			return _type.GetField(name, bindingAttr);
		}

		public FieldInfo[] GetFields(BindingFlags bindingAttr)
		{
			return _type.GetFields(bindingAttr);
		}

		public PropertyInfo GetProperty(string name, BindingFlags bindingAttr)
		{
			return _type.GetProperty(name, bindingAttr);
		}

		public PropertyInfo GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers)
		{
			return _type.GetProperty(name, bindingAttr, binder, returnType, types, modifiers);
		}

		public PropertyInfo[] GetProperties(BindingFlags bindingAttr)
		{
			return _type.GetProperties(bindingAttr);
		}

		public MemberInfo[] GetMember(string name, BindingFlags bindingAttr)
		{
			return _type.GetMember(name, bindingAttr);
		}

		public MemberInfo[] GetMembers(BindingFlags bindingAttr)
		{
			return _type.GetMembers(bindingAttr);
		}

		public object InvokeMember(string name, BindingFlags invokeAttr, Binder binder, object target, object[] args, ParameterModifier[] modifiers, CultureInfo culture, string[] namedParameters)
		{
			if (name == "[DISPID=0]")
			{
				OnHtmlEvent(args == null ? null : args.Length == 0 ? null : args[0]);
				return null;
			}
			return _type.InvokeMember(name, invokeAttr, binder, target, args, modifiers, culture, namedParameters);
		}

		public Type UnderlyingSystemType
		{
			get { return _type.UnderlyingSystemType; }
		}

		public bool Equals(EventProxy other)
		{
			if (ReferenceEquals(null, other)) return false;
			if (ReferenceEquals(this, other)) return true;
			return Equals(other._eventName, _eventName) && Equals(other._target, _target);
		}

		public override bool Equals(object obj)
		{
			if (ReferenceEquals(null, obj)) return false;
			if (ReferenceEquals(this, obj)) return true;
			return obj.GetType() == typeof (EventProxy) && Equals((EventProxy) obj);
		}

		public override int GetHashCode()
		{
			unchecked
			{
				return (_eventName.GetHashCode()*397) ^ _target.GetHashCode();
			}
		}

		public static bool operator ==(EventProxy left, EventProxy right)
		{
			return Equals(left, right);
		}

		public static bool operator !=(EventProxy left, EventProxy right)
		{
			return !Equals(left, right);
		}
	}
}
