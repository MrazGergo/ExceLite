using System;

namespace ExceLite.Exceptions
{
    public class NoValidPropertyException : Exception
    {
        public Type Type { get; }

        public NoValidPropertyException(Type type) : base($"The type {type.Name} does not have any public properties.")
        {
            Type = type;
        }
    }
}
