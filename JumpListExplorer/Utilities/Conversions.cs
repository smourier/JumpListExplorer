using System;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace JumpListExplorer.Utilities
{
    public static class Conversions
    {
        public static bool EqualsIgnoreCase(this string? thisString, string? text, bool trim = true)
        {
            if (trim)
            {
                thisString = thisString.Nullify();
                text = text.Nullify();
            }

            if (thisString == null)
                return text == null;

            if (text == null)
                return false;

            if (thisString.Length != text.Length)
                return false;

            return string.Compare(thisString, text, StringComparison.OrdinalIgnoreCase) == 0;
        }

        public static string? Nullify(this string? text)
        {
            if (text == null)
                return null;

            if (string.IsNullOrWhiteSpace(text))
                return null;

            var t = text.Trim();
            return t.Length == 0 ? null : t;
        }

        public static object? ChangeType(object? input, Type conversionType, object? defaultValue = null, IFormatProvider? provider = null)
        {
            if (!TryChangeType(input, conversionType, provider, out object? value))
                return defaultValue;

            return value;
        }

        public static T? ChangeType<T>(object? input, T? defaultValue = default, IFormatProvider? provider = null)
        {
            if (!TryChangeType(input, provider, out T? value))
                return defaultValue;

            return value;
        }

        public static bool TryChangeType<T>(object? input, out T? value) => TryChangeType(input, null, out value);
        public static bool TryChangeType<T>(object? input, IFormatProvider? provider, out T? value)
        {
            if (!TryChangeType(input, typeof(T), provider, out object? tvalue))
            {
                value = default;
                return false;
            }

            value = (T?)tvalue;
            return true;
        }

        public static bool TryChangeType(object? input, Type conversionType, out object? value) => TryChangeType(input, conversionType, null, out value);
        public static bool TryChangeType(object? input, Type conversionType, IFormatProvider? provider, out object? value)
        {
            ArgumentNullException.ThrowIfNull(conversionType);
            if (conversionType == typeof(object))
            {
                value = input;
                return true;
            }

            value = conversionType.IsValueType ? RuntimeHelpers.GetUninitializedObject(conversionType) : null;
            if (input == null)
                return !conversionType.IsValueType;

            var inputType = input.GetType();
            if (conversionType.IsAssignableFrom(inputType))
            {
                value = input;
                return true;
            }

            if (conversionType.IsEnum)
                return TryParseEnum(conversionType, input, out value);

            if (conversionType == typeof(Guid))
            {
                var svalue = string.Format(provider, "{0}", input).Nullify();
                if (svalue != null && Guid.TryParse(svalue, out Guid guid))
                {
                    value = guid;
                    return true;
                }
                return false;
            }

            if (conversionType == typeof(Type))
            {
                var typeName = string.Format(provider, "{0}", input).Nullify();
                if (typeName == null)
                    return false;

                var type = Type.GetType(typeName, false);
                if (type == null)
                    return false;

                value = type;
                return true;
            }

            if (conversionType == typeof(IntPtr))
            {
                if (IntPtr.Size == 8)
                {
                    if (TryChangeType(input, provider, out long l))
                    {
                        value = new IntPtr(l);
                        return true;
                    }
                }
                else if (TryChangeType(input, provider, out int i2))
                {
                    value = new IntPtr(i2);
                    return true;
                }
                return false;
            }

            if (conversionType == typeof(int))
            {
                if (inputType == typeof(uint))
                {
                    value = unchecked((int)(uint)input);
                    return true;
                }

                if (inputType == typeof(ulong))
                {
                    value = unchecked((int)(ulong)input);
                    return true;
                }

                if (inputType == typeof(ushort))
                {
                    value = unchecked((int)(ushort)input);
                    return true;
                }

                if (inputType == typeof(byte))
                {
                    value = unchecked((int)(byte)input);
                    return true;
                }

                if (input is string s)
                {
                    if (int.TryParse(s, NumberStyles.Any, provider, out var si))
                    {
                        value = si;
                        return true;
                    }
                    return false;
                }
            }

            if (conversionType == typeof(long))
            {
                if (inputType == typeof(uint))
                {
                    value = unchecked((long)(uint)input);
                    return true;
                }

                if (inputType == typeof(ulong))
                {
                    value = unchecked((long)(ulong)input);
                    return true;
                }

                if (inputType == typeof(ushort))
                {
                    value = unchecked((long)(ushort)input);
                    return true;
                }

                if (inputType == typeof(byte))
                {
                    value = unchecked((long)(byte)input);
                    return true;
                }

                if (input is string s)
                {
                    if (long.TryParse(s, NumberStyles.Any, provider, out var sl))
                    {
                        value = sl;
                        return true;
                    }
                    return false;
                }
            }

            if (conversionType == typeof(short))
            {
                if (inputType == typeof(uint))
                {
                    value = unchecked((short)(uint)input);
                    return true;
                }

                if (inputType == typeof(ulong))
                {
                    value = unchecked((short)(ulong)input);
                    return true;
                }

                if (inputType == typeof(ushort))
                {
                    value = unchecked((short)(ushort)input);
                    return true;
                }

                if (inputType == typeof(byte))
                {
                    value = unchecked((short)(byte)input);
                    return true;
                }

                if (input is string s)
                {
                    if (short.TryParse(s, NumberStyles.Any, provider, out var ss))
                    {
                        value = ss;
                        return true;
                    }
                    return false;
                }
            }

            if (conversionType == typeof(sbyte))
            {
                if (inputType == typeof(uint))
                {
                    value = unchecked((sbyte)(uint)input);
                    return true;
                }

                if (inputType == typeof(ulong))
                {
                    value = unchecked((sbyte)(ulong)input);
                    return true;
                }

                if (inputType == typeof(ushort))
                {
                    value = unchecked((sbyte)(ushort)input);
                    return true;
                }

                if (inputType == typeof(byte))
                {
                    value = unchecked((sbyte)(byte)input);
                    return true;
                }

                if (input is string s)
                {
                    if (sbyte.TryParse(s, NumberStyles.Any, provider, out var sb))
                    {
                        value = sb;
                        return true;
                    }
                    return false;
                }
            }

            if (conversionType == typeof(uint))
            {
                if (inputType == typeof(int))
                {
                    value = unchecked((uint)(int)input);
                    return true;
                }

                if (inputType == typeof(long))
                {
                    value = unchecked((uint)(long)input);
                    return true;
                }

                if (inputType == typeof(short))
                {
                    value = unchecked((uint)(short)input);
                    return true;
                }

                if (inputType == typeof(sbyte))
                {
                    value = unchecked((uint)(sbyte)input);
                    return true;
                }

                if (input is string s)
                {
                    if (uint.TryParse(s, NumberStyles.Any, provider, out var ui))
                    {
                        value = ui;
                        return true;
                    }
                    return false;
                }
            }

            if (conversionType == typeof(ulong))
            {
                if (inputType == typeof(int))
                {
                    value = unchecked((ulong)(int)input);
                    return true;
                }

                if (inputType == typeof(long))
                {
                    value = unchecked((ulong)(long)input);
                    return true;
                }

                if (inputType == typeof(short))
                {
                    value = unchecked((ulong)(short)input);
                    return true;
                }

                if (inputType == typeof(sbyte))
                {
                    value = unchecked((ulong)(sbyte)input);
                    return true;
                }

                if (input is string s)
                {
                    if (ulong.TryParse(s, NumberStyles.Any, provider, out var ul))
                    {
                        value = ul;
                        return true;
                    }
                    return false;
                }
            }

            if (conversionType == typeof(ushort))
            {
                if (inputType == typeof(int))
                {
                    value = unchecked((ushort)(int)input);
                    return true;
                }

                if (inputType == typeof(long))
                {
                    value = unchecked((ushort)(long)input);
                    return true;
                }

                if (inputType == typeof(short))
                {
                    value = unchecked((ushort)(short)input);
                    return true;
                }

                if (inputType == typeof(sbyte))
                {
                    value = unchecked((ushort)(sbyte)input);
                    return true;
                }

                if (input is string s)
                {
                    if (ushort.TryParse(s, NumberStyles.Any, provider, out var us))
                    {
                        value = us;
                        return true;
                    }
                    return false;
                }
            }

            if (conversionType == typeof(byte))
            {
                if (inputType == typeof(int))
                {
                    value = unchecked((byte)(int)input);
                    return true;
                }

                if (inputType == typeof(long))
                {
                    value = unchecked((byte)(long)input);
                    return true;
                }

                if (inputType == typeof(short))
                {
                    value = unchecked((byte)(short)input);
                    return true;
                }

                if (inputType == typeof(sbyte))
                {
                    value = unchecked((byte)(sbyte)input);
                    return true;
                }

                if (input is string s)
                {
                    if (byte.TryParse(s, NumberStyles.Any, provider, out var b))
                    {
                        value = b;
                        return true;
                    }
                    return false;
                }
            }

            if (conversionType == typeof(float))
            {
                if (input is string s)
                {
                    if (float.TryParse(s, NumberStyles.Any, provider, out var fl))
                    {
                        value = fl;
                        return true;
                    }
                    return false;
                }
            }

            if (conversionType == typeof(double))
            {
                if (input is string s)
                {
                    if (double.TryParse(s, NumberStyles.Any, provider, out var dbl3))
                    {
                        value = dbl3;
                        return true;
                    }
                    return false;
                }
            }

            if (conversionType == typeof(decimal))
            {
                if (input is string s)
                {
                    if (decimal.TryParse(s, NumberStyles.Any, provider, out var dec))
                    {
                        value = dec;
                        return true;
                    }
                    return false;
                }
            }

            if (conversionType == typeof(DateTime) && input is double dbl)
            {
                try
                {
                    value = DateTime.FromOADate(dbl);
                    return true;
                }
                catch
                {
                    value = DateTime.MinValue;
                    return false;
                }
            }

            if (conversionType == typeof(DateTimeOffset) && input is double dbl2)
            {
                try
                {
                    value = new DateTimeOffset(DateTime.FromOADate(dbl2));
                    return true;
                }
                catch
                {
                    value = DateTimeOffset.MinValue;
                    return false;
                }
            }

            if (conversionType == typeof(bool) && TryChangeType<long>(input, out var i))
            {
                value = i != 0;
                return true;
            }

            var nullable = conversionType.IsGenericType && conversionType.GetGenericTypeDefinition() == typeof(Nullable<>);
            if (nullable)
            {
                if (string.Empty.Equals(input))
                {
                    value = null;
                    return true;
                }

                var type = conversionType.GetGenericArguments()[0];
                if (TryChangeType(input, type, provider, out var vtValue))
                {
                    var nullableType = typeof(Nullable<>).MakeGenericType(type);
                    value = Activator.CreateInstance(nullableType, vtValue);
                    return true;
                }

                value = null;
                return false;
            }

            if (input is IConvertible convertible)
            {
                try
                {
                    value = convertible.ToType(conversionType, provider);
                    return true;
                }
                catch
                {
                    // do nothing
                    return false;
                }
            }

            return false;
        }

        public static ulong EnumToUInt64(object value)
        {
            ArgumentNullException.ThrowIfNull(value);
            var typeCode = Convert.GetTypeCode(value);
            switch (typeCode)
            {
                case TypeCode.SByte:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                    return (ulong)Convert.ToInt64(value, CultureInfo.InvariantCulture);

                case TypeCode.Byte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    return Convert.ToUInt64(value, CultureInfo.InvariantCulture);

                case TypeCode.String:
                default:
                    return ChangeType<ulong>(value, 0, CultureInfo.InvariantCulture);
            }
        }

        public static bool IsFlagsEnum(Type enumType)
        {
            ArgumentNullException.ThrowIfNull(enumType);
            if (!enumType.IsEnum)
                throw new ArgumentException(null, nameof(enumType));

            return enumType.IsDefined(typeof(FlagsAttribute), true);
        }

        public static int GetEnumMaxPower(Type enumType)
        {
            ArgumentNullException.ThrowIfNull(enumType);
            if (!enumType.IsEnum)
                throw new ArgumentException(null, nameof(enumType));

            var type = Enum.GetUnderlyingType(enumType);
            return GetEnumUnderlyingTypeMaxPower(type);
        }

        public static int GetEnumUnderlyingTypeMaxPower(Type underlyingType)
        {
            ArgumentNullException.ThrowIfNull(underlyingType);
            if (underlyingType == typeof(long) ||
                underlyingType == typeof(ulong))
                return 64;

            if (underlyingType == typeof(int) ||
                underlyingType == typeof(uint))
                return 32;

            if (underlyingType == typeof(short) ||
                underlyingType == typeof(ushort))
                return 16;

            if (underlyingType == typeof(byte) ||
                underlyingType == typeof(sbyte))
                return 8;

            throw new ArgumentException(null, nameof(underlyingType));
        }

        public static object EnumToObject(Type enumType, object value)
        {
            ArgumentNullException.ThrowIfNull(enumType);
            ArgumentNullException.ThrowIfNull(value);
            if (!enumType.IsEnum)
                throw new ArgumentException(null, nameof(enumType));

            var underlyingType = Enum.GetUnderlyingType(enumType);
            if (underlyingType == typeof(long))
                return Enum.ToObject(enumType, ChangeType<long>(value));

            if (underlyingType == typeof(ulong))
                return Enum.ToObject(enumType, ChangeType<ulong>(value));

            if (underlyingType == typeof(int))
                return Enum.ToObject(enumType, ChangeType<int>(value));

            if (underlyingType == typeof(uint))
                return Enum.ToObject(enumType, ChangeType<uint>(value));

            if (underlyingType == typeof(short))
                return Enum.ToObject(enumType, ChangeType<short>(value));

            if (underlyingType == typeof(ushort))
                return Enum.ToObject(enumType, ChangeType<ushort>(value));

            if (underlyingType == typeof(byte))
                return Enum.ToObject(enumType, ChangeType<byte>(value));

            if (underlyingType == typeof(sbyte))
                return Enum.ToObject(enumType, ChangeType<sbyte>(value));

            throw new ArgumentException(null, nameof(enumType));
        }

        private static readonly char[] _enumSeparators = new char[] { ',', ';', '+', '|', ' ' };

        public static object ToEnum(string text, Type enumType) { TryParseEnum(enumType, text, out object value); return value; }
        public static bool TryParseEnum(Type type, object input, out object value)
        {
            ArgumentNullException.ThrowIfNull(type);
            if (!type.IsEnum)
                throw new ArgumentException(null, nameof(type));

            if (input == null)
            {
                value = RuntimeHelpers.GetUninitializedObject(type);
                return false;
            }

            var stringInput = string.Format(CultureInfo.InvariantCulture, "{0}", input);
            stringInput = stringInput.Nullify();
            if (stringInput == null)
            {
                value = RuntimeHelpers.GetUninitializedObject(type);
                return false;
            }

            if (stringInput.StartsWith("0x", StringComparison.OrdinalIgnoreCase) && ulong.TryParse(stringInput.AsSpan(2), NumberStyles.HexNumber, null, out var ulx))
            {
                value = ToEnum(ulx.ToString(CultureInfo.InvariantCulture), type);
                return true;
            }

            var names = Enum.GetNames(type);
            if (names.Length == 0)
            {
                value = RuntimeHelpers.GetUninitializedObject(type);
                return false;
            }

            var values = Enum.GetValues(type);
            // some enums like System.CodeDom.MemberAttributes *are* flags but are not declared with Flags...
            if (!type.IsDefined(typeof(FlagsAttribute), true) && stringInput.IndexOfAny(_enumSeparators) < 0)
                return StringToEnum(type, names, values, stringInput, out value);

            // multi value enum
            var tokens = stringInput.Split(_enumSeparators, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0)
            {
                value = RuntimeHelpers.GetUninitializedObject(type);
                return false;
            }

            ulong ul = 0;
            foreach (var tok in tokens)
            {
                var token = tok.Nullify(); // NOTE: we don't consider empty tokens as errors
                if (token == null)
                    continue;

                if (!StringToEnum(type, names, values, token, out object tokenValue))
                {
                    value = RuntimeHelpers.GetUninitializedObject(type);
                    return false;
                }

                ulong tokenUl;
                switch (Convert.GetTypeCode(tokenValue))
                {
                    case TypeCode.Int16:
                    case TypeCode.Int32:
                    case TypeCode.Int64:
                    case TypeCode.SByte:
                        tokenUl = (ulong)Convert.ToInt64(tokenValue, CultureInfo.InvariantCulture);
                        break;

                    default:
                        tokenUl = Convert.ToUInt64(tokenValue, CultureInfo.InvariantCulture);
                        break;
                }

                ul |= tokenUl;
            }
            value = Enum.ToObject(type, ul);
            return true;
        }

        private static bool StringToEnum(Type type, string[] names, Array values, string input, out object value)
        {
            for (var i = 0; i < names.Length; i++)
            {
                if (names[i].EqualsIgnoreCase(input))
                {
                    value = values.GetValue(i)!;
                    return true;
                }
            }

            for (var i = 0; i < values.GetLength(0); i++)
            {
                var valuei = values.GetValue(i)!;
                if (input.Length > 0 && input[0] == '-')
                {
                    var ul = (long)EnumToUInt64(valuei);
                    if (ul.ToString().EqualsIgnoreCase(input))
                    {
                        value = valuei;
                        return true;
                    }
                }
                else
                {
                    var ul = EnumToUInt64(valuei);
                    if (ul.ToString().EqualsIgnoreCase(input))
                    {
                        value = valuei;
                        return true;
                    }
                }
            }

            if (char.IsDigit(input[0]) || input[0] == '-' || input[0] == '+')
            {
                var obj = EnumToObject(type, input);
                if (obj == null)
                {
                    value = RuntimeHelpers.GetUninitializedObject(type);
                    return false;
                }
                value = obj;
                return true;
            }

            value = RuntimeHelpers.GetUninitializedObject(type);
            return false;
        }
    }
}
