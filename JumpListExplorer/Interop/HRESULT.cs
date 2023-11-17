using System;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace JumpListExplorer.Interop
{
    [StructLayout(LayoutKind.Sequential)]
    internal readonly struct HRESULT : IEquatable<HRESULT>, IFormattable
    {
        private static readonly ConcurrentDictionary<uint, string?> _names = new();

        public HRESULT(uint value)
        {
            Value = value;
        }

        public HRESULT(int value)
        {
            Value = (uint)value;
        }

        public HRESULT(HRESULTS value)
            : this((uint)value)
        {
        }

        public uint Value { get; }
        public int IValue => (int)Value;
        public readonly string Name => ToString("n", null);
        public readonly bool IsError => Value < 0;
        public readonly bool IsSuccess => Value >= 0;
        public readonly bool IsOk => Value == (int)HRESULTS.S_OK;
        public readonly bool IsFalse => Value == (int)HRESULTS.S_FALSE;

        public uint ThrowOnError(bool throwOnError = true)
        {
            if (!throwOnError || Value == 0)
                return Value;

            if (Value == (uint)HRESULTS.DISP_E_EXCEPTION)
            {
                var error = ComError.GetError();
                if (error != null)
                    throw error;
            }

            if (Value == (uint)HRESULTS.STG_E_FILENOTFOUND)
                throw new FileNotFoundException();

            if (Value == (uint)HRESULTS.STG_E_PATHNOTFOUND)
                throw new DirectoryNotFoundException();

            if (Value < 0)
                throw new Win32Exception((int)Value);

            return Value;
        }

        public readonly HRESULT ToHRESULT() => new(Value);
        public readonly int ToInt32() => (int)Value;
        public readonly uint ToUInt32() => Value;
        public readonly HRESULTS ToHRESULTS() => (HRESULTS)Value;

        public override readonly bool Equals(object? obj) => Value.Equals(obj);
        public override readonly int GetHashCode() => Value.GetHashCode();
        public readonly bool Equals(HRESULT other) => Value.Equals(other.Value);

        public override readonly string ToString() => ToString(null, null);
        public readonly string ToString(string? format, IFormatProvider? formatProvider)
        {
            switch (format?.ToUpperInvariant())
            {
                case "I":
                    return Value.ToString(CultureInfo.InvariantCulture);

                case "N":
                    if (!_names.TryGetValue(Value, out var text))
                    {
                        var value = Value;
                        text = typeof(HRESULTS).GetFields(BindingFlags.Static | BindingFlags.Public).FirstOrDefault(f => (int)(HRESULTS)f.GetValue(null)! == value)?.Name;
                        _names[Value] = text;
                    }
                    return text ?? string.Empty;

                case "U":
                    return Value.ToString(CultureInfo.InvariantCulture);

                case "X":
                    return "0x" + Value.ToString("X8", CultureInfo.InvariantCulture);

                default:
                    var name = ToString("n", formatProvider);
                    if (name != null)
                        return name + " (0x" + Value.ToString("X8", CultureInfo.InvariantCulture) + ")";

                    return "0x" + Value.ToString("X8", CultureInfo.InvariantCulture);
            }
        }

        public static bool operator ==(HRESULT left, HRESULT right) => left.Value == right.Value;
        public static bool operator !=(HRESULT left, HRESULT right) => left.Value != right.Value;
        public static implicit operator HRESULT(int value) => new(value);
        public static implicit operator HRESULT(uint result) => new(result);
        public static implicit operator HRESULT(HRESULTS result) => new(result);
        public static explicit operator uint(HRESULT hr) => hr.Value;
        public static explicit operator int(HRESULT hr) => hr.IValue;
        public static explicit operator HRESULTS(HRESULT hr) => (HRESULTS)hr.Value;
    }
}
