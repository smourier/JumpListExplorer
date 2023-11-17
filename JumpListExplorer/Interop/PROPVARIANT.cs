using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace JumpListExplorer.Interop
{
    [StructLayout(LayoutKind.Explicit)]
    public sealed class PROPVARIANT : IDisposable
    {
#pragma warning disable IDE0044 // Add readonly modifier
        [FieldOffset(0)]
        private VARTYPE _vt;

        [FieldOffset(8)]
        private IntPtr _ptr;

        [FieldOffset(8)]
        private int _int32;

        [FieldOffset(8)]
        private uint _uint32;

        [FieldOffset(8)]
        private byte _byte;

        [FieldOffset(8)]
        private sbyte _sbyte;

        [FieldOffset(8)]
        private short _int16;

        [FieldOffset(8)]
        private ushort _uint16;

        [FieldOffset(8)]
        private long _int64;

        [FieldOffset(8)]
        private ulong _uint64;

        [FieldOffset(8)]
        private double _double;

        [FieldOffset(8)]
        private float _single;

        [FieldOffset(8)]
        private short _boolean;

        [FieldOffset(8)]
        private System.Runtime.InteropServices.ComTypes.FILETIME _filetime;

        [FieldOffset(8)]
        private PROPARRAY _ca;

        [FieldOffset(0)]
        private decimal _decimal;
#pragma warning restore IDE0044 // Add readonly modifier

        [StructLayout(LayoutKind.Sequential)]
        private struct PROPARRAY
        {
            public int cElems;
            public IntPtr pElems;
        }

        public PROPVARIANT()
        {
            // it's a VT_EMPTY
        }

        private void ConstructBlob(byte[] bytes)
        {
            _ca.cElems = bytes.Length;
            _ca.pElems = Marshal.AllocCoTaskMem(bytes.Length);
            Marshal.Copy(bytes, 0, _ca.pElems, bytes.Length);
            _vt = VARTYPE.VT_BLOB;
        }

        private void ConstructArray(Array array, VARTYPE? type = null)
        {
            // special case for bools which are shorts...
            if (array is bool[] bools)
            {
                var shorts = new short[bools.Length];
                for (var i = 0; i < bools.Length; i++)
                {
                    shorts[i] = bools[i] ? (short)-1 : (short)0;
                }
                ConstructVector(shorts, typeof(short), VARTYPE.VT_BOOL);
                return;
            }

            var elementType = array.GetType().GetElementType()!;
            if (type == VARTYPE.VT_BLOB)
            {
                if (array is not byte[] bytes)
                    throw new ArgumentException("Property type " + type + " is only supported for arrays of bytes.", nameof(type));

                ConstructBlob(bytes);
                return;
            }

            ConstructVector(array, elementType, FromType(elementType, type));
        }

        private void ConstructVector(Array array, Type type, VARTYPE vt)
        {
            _vt = vt | VARTYPE.VT_VECTOR;
            if (array.Length > 0)
            {
                int size;
                if (type == typeof(string))
                {
                    size = IntPtr.Size;
                }
                else
                {
                    size = Marshal.SizeOf(type);
                }

                size *= array.Length;
                var ptr = Marshal.AllocCoTaskMem(size);
                _ca.cElems = array.Length;
                _ca.pElems = ptr;

                if (type == typeof(string))
                {
                    for (var i = 0; i < array.Length; i++)
                    {
                        var str = MarshalString((string?)array.GetValue(i), vt);
                        Marshal.WriteIntPtr(ptr, IntPtr.Size * i, str);
                    }
                }
                else
                {
                    CopyMemory(ptr, Marshal.UnsafeAddrOfPinnedArrayElement(array, 0), (IntPtr)size);
                }
            }
        }

        private static void Using(object resource, Action action)
        {
            try
            {
                action();
            }
            finally
            {
                (resource as IDisposable)?.Dispose();
            }
        }

        private static int GetCount(IEnumerable enumerable)
        {
            if (enumerable is ICollection col)
                return col.Count;

            var count = 0;
            var e = enumerable.GetEnumerator();
            Using(e, () =>
            {
                while (e.MoveNext())
                {
                    count++;
                }
            });
            return count;
        }

        private static Type? GetElementType(Type collectionType)
        {
            foreach (var iface in collectionType.GetInterfaces())
            {
                if (!iface.IsGenericType)
                    continue;

                if (iface.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                    return iface.GetGenericArguments()[0];

                if (iface.GetGenericTypeDefinition() == typeof(ICollection<>))
                    return iface.GetGenericArguments()[0];

                if (iface.GetGenericTypeDefinition() == typeof(IList<>))
                    return iface.GetGenericArguments()[0];
            }
            return null;
        }

        private static Type? GetElementType(IEnumerable enumerable)
        {
            var elementType = GetElementType(enumerable.GetType());
            if (elementType != null)
                return elementType;

            foreach (var obj in enumerable)
            {
                return obj.GetType();
            }
            return null;
        }

        private void ConstructEnumerable(IEnumerable enumerable, VARTYPE? type = null)
        {
            var elementType = GetElementType(enumerable);
            if (elementType == null)
                throw new ArgumentException("Enumerable type '" + enumerable.GetType().FullName + "' is not supported.", nameof(enumerable));

            var count = GetCount(enumerable);
            var array = Array.CreateInstance(elementType, count);
            var i = 0;
            foreach (var obj in enumerable)
            {
                array.SetValue(obj, i++);
            }
            ConstructArray(array, type);
        }

        private static Type FromType(VARTYPE type)
        {
            switch (type)
            {
                case VARTYPE.VT_I1:
                    return typeof(sbyte);

                case VARTYPE.VT_UI1:
                    return typeof(byte);

                case VARTYPE.VT_I2:
                    return typeof(short);

                case VARTYPE.VT_UI2:
                    return typeof(ushort);

                case VARTYPE.VT_UI4:
                case VARTYPE.VT_UINT:
                    return typeof(uint);

                case VARTYPE.VT_I8:
                    return typeof(long);

                case VARTYPE.VT_UI8:
                    return typeof(ulong);

                case VARTYPE.VT_R4:
                    return typeof(float);

                case VARTYPE.VT_R8:
                    return typeof(double);

                case VARTYPE.VT_BOOL:
                    return typeof(bool);

                case VARTYPE.VT_I4:
                case VARTYPE.VT_INT:
                case VARTYPE.VT_ERROR:
                    return typeof(int);

                case VARTYPE.VT_DATE:
                    return typeof(DateTime);

                case VARTYPE.VT_FILETIME:
                    return typeof(System.Runtime.InteropServices.ComTypes.FILETIME);

                case VARTYPE.VT_BLOB:
                    return typeof(byte[]);

                case VARTYPE.VT_CLSID:
                    return typeof(Guid);

                case VARTYPE.VT_BSTR:
                case VARTYPE.VT_LPSTR:
                case VARTYPE.VT_LPWSTR:
                    return typeof(string);

                case VARTYPE.VT_UNKNOWN:
                case VARTYPE.VT_DISPATCH:
                    return typeof(object);

                case VARTYPE.VT_CY:
                case VARTYPE.VT_DECIMAL:
                    return typeof(decimal);

                default:
                    throw new ArgumentException("Property type " + type + " is not supported.", nameof(type));
            }
        }

        private static VARTYPE FromType(Type type, VARTYPE? vt)
        {
            if (type == null)
                return VARTYPE.VT_NULL;

            var tc = Type.GetTypeCode(type);
            switch (tc)
            {
                case TypeCode.Boolean:
                    return VARTYPE.VT_BOOL;

                case TypeCode.Byte:
                    return VARTYPE.VT_UI1;

                case TypeCode.Char:
                    return VARTYPE.VT_LPWSTR;

                case TypeCode.DateTime:
                    return VARTYPE.VT_FILETIME;

                case TypeCode.DBNull:
                    return VARTYPE.VT_NULL;

                case TypeCode.Decimal:
                    return VARTYPE.VT_DECIMAL;

                case TypeCode.Double:
                    return VARTYPE.VT_R8;

                case TypeCode.Empty:
                    return VARTYPE.VT_EMPTY;

                case TypeCode.Int16:
                    return VARTYPE.VT_I2;

                case TypeCode.Int32:
                    return VARTYPE.VT_I4;

                case TypeCode.Int64:
                    return VARTYPE.VT_I8;

                case TypeCode.SByte:
                    return VARTYPE.VT_I1;

                case TypeCode.Single:
                    return VARTYPE.VT_R4;

                case TypeCode.String:
                    if (!vt.HasValue)
                        return VARTYPE.VT_LPWSTR;

                    if (vt != VARTYPE.VT_LPSTR && vt != VARTYPE.VT_BSTR && vt != VARTYPE.VT_LPWSTR)
                        throw new ArgumentException("Property type " + vt + " is not supported for string.", nameof(type));

                    return vt.Value;

                case TypeCode.UInt16:
                    return VARTYPE.VT_UI2;

                case TypeCode.UInt32:
                    return VARTYPE.VT_UI4;

                case TypeCode.UInt64:
                    return VARTYPE.VT_UI8;

                // case TypeCode.Object:
                default:
                    if (type == typeof(Guid))
                        return VARTYPE.VT_CLSID;

                    if (type == typeof(System.Runtime.InteropServices.ComTypes.FILETIME))
                        return VARTYPE.VT_FILETIME;

                    if (type == typeof(byte))
                    {
                        if (!vt.HasValue)
                            return VARTYPE.VT_UI1 | VARTYPE.VT_VECTOR;

                        if (vt != VARTYPE.VT_BLOB && vt != (VARTYPE.VT_UI1 | VARTYPE.VT_VECTOR))
                            throw new ArgumentException("Property type " + vt + " is not supported for array of bytes.", nameof(type));

                        return vt.Value;
                    }

                    throw new ArgumentException("Value of type '" + type.FullName + "' is not supported.", nameof(type));
            }
        }

        public PROPVARIANT(object? value, VARTYPE? type = null)
        {
            if (value is PROPVARIANT pv)
            {
                value = pv.Value;
            }

            if (value == null)
            {
                _vt = VARTYPE.VT_NULL;
                return;
            }

            if (Marshal.IsComObject(value))
            {
                _ptr = Marshal.GetIUnknownForObject(value);
                _vt = VARTYPE.VT_UNKNOWN;
                return;
            }

            if (value is char[] chars)
            {
                value = new string(chars);
            }

            if (value is char[][] charray)
            {
                var strings = new string[charray.GetLength(0)];
                for (var i = 0; i < charray.Length; i++)
                {
                    strings[i] = new string(charray[i]);
                }
                value = strings;
            }

            if (value is Array array)
            {
                ConstructArray(array, type);
                return;
            }

            if (value is not string && value is IEnumerable enumerable)
            {
                ConstructEnumerable(enumerable, type);
                return;
            }

            var tc = Type.GetTypeCode(value.GetType());
            switch (tc)
            {
                case TypeCode.Boolean:
                    _boolean = (bool)value ? (short)-1 : (short)0;
                    break;

                case TypeCode.Byte:
                    _byte = (byte)value;
                    break;

                case TypeCode.Char:
                    chars = new[] { (char)value };
                    _ptr = MarshalString(new string(chars), FromType(typeof(string), type));
                    break;

                case TypeCode.DateTime:
                    var ft = ToPositiveFileTime((DateTime)value);
                    if (ft == 0)
                        break; // stay empty

                    InitPropVariantFromFileTime(ref ft, this);
                    break;

                case TypeCode.Empty:
                case TypeCode.DBNull:
                    break;

                case TypeCode.Decimal:
                    _decimal = (decimal)value;
                    break;

                case TypeCode.Double:
                    _double = (double)value;
                    break;

                case TypeCode.Int16:
                    _int16 = (short)value;
                    break;

                case TypeCode.Int32:
                    _int32 = (int)value;
                    break;

                case TypeCode.Int64:
                    _int64 = (long)value;
                    break;

                case TypeCode.SByte:
                    _sbyte = (sbyte)value;
                    break;

                case TypeCode.Single:
                    _single = (float)value;
                    break;

                case TypeCode.String:
                    _ptr = MarshalString((string)value, FromType(typeof(string), type));
                    break;

                case TypeCode.UInt16:
                    _uint16 = (ushort)value;
                    break;

                case TypeCode.UInt32:
                    _uint32 = (uint)value;
                    break;

                case TypeCode.UInt64:
                    _uint64 = (ulong)value;
                    break;

                //case TypeCode.Object:
                default:
                    if (value is Guid guid)
                    {
                        _ptr = Marshal.AllocCoTaskMem(16);
                        Marshal.Copy(guid.ToByteArray(), 0, _ptr, 16);
                        break;
                    }

                    if (value is System.Runtime.InteropServices.ComTypes.FILETIME filetime)
                    {
                        _filetime = filetime;
                        break;
                    }
                    throw new ArgumentException("Value of type '" + value.GetType().FullName + "' is not supported.", nameof(value));
            }

            _vt = FromType(value.GetType(), type);
        }

        public VARTYPE VarType { get => _vt; set => _vt = value; }
        public object? Value
        {
            get
            {
                switch (_vt)
                {
                    case VARTYPE.VT_EMPTY:
                    case VARTYPE.VT_NULL: // DbNull
                        return null;

                    case VARTYPE.VT_I1:
                        return _sbyte;

                    case VARTYPE.VT_UI1:
                        return _byte;

                    case VARTYPE.VT_I2:
                        return _int16;

                    case VARTYPE.VT_UI2:
                        return _uint16;

                    case VARTYPE.VT_I4:
                    case VARTYPE.VT_INT:
                        return _int32;

                    case VARTYPE.VT_UI4:
                    case VARTYPE.VT_UINT:
                        return _uint32;

                    case VARTYPE.VT_I8:
                        return _int64;

                    case VARTYPE.VT_UI8:
                        return _uint64;

                    case VARTYPE.VT_R4:
                        return _single;

                    case VARTYPE.VT_R8:
                        return _double;

                    case VARTYPE.VT_BOOL:
                        return _int32 != 0;

                    case VARTYPE.VT_ERROR:
                        return _int64;

                    case VARTYPE.VT_CY:
                        return _decimal;

                    case VARTYPE.VT_DATE:
                        return DateTime.FromOADate(_double);

                    case VARTYPE.VT_FILETIME:
                        return DateTime.FromFileTime(_int64);

                    case VARTYPE.VT_BSTR:
                        return Marshal.PtrToStringBSTR(_ptr);

                    case VARTYPE.VT_BLOB:
                        var blob = new byte[_ca.cElems];
                        Marshal.Copy(_ca.pElems, blob, 0, _int32);
                        return blob;

                    case VARTYPE.VT_CLSID:
                        var guid = new byte[16];
                        Marshal.Copy(_ptr, guid, 0, guid.Length);
                        return new Guid(guid);

                    case VARTYPE.VT_LPSTR:
                        return Marshal.PtrToStringAnsi(_ptr);

                    case VARTYPE.VT_LPWSTR:
                        return Marshal.PtrToStringUni(_ptr);

                    case VARTYPE.VT_UNKNOWN:
                    case VARTYPE.VT_DISPATCH:
                        return Marshal.GetObjectForIUnknown(_ptr);

                    case VARTYPE.VT_DECIMAL:
                        return _decimal;

                    default:
                        if ((_vt & VARTYPE.VT_VECTOR) == VARTYPE.VT_VECTOR)
                        {
                            var et = _vt & ~VARTYPE.VT_VECTOR;
                            if (TryGetVectorValue(et, out var vector))
                                return vector;
                        }
                        throw new NotSupportedException("Value of property type " + _vt + " is not supported.");
                }
            }
        }

        ~PROPVARIANT() => Dispose();
        public void Dispose()
        {
            _ = PropVariantClear(this);
            GC.SuppressFinalize(this);
        }

        private static IntPtr MarshalString(string? str, VARTYPE vt)
        {
            switch (vt)
            {
                case VARTYPE.VT_LPWSTR:
                    return Marshal.StringToCoTaskMemUni(str);

                case VARTYPE.VT_BSTR:
                    return Marshal.StringToBSTR(str);

                case VARTYPE.VT_LPSTR:
                    return Marshal.StringToCoTaskMemAnsi(str);

                default:
                    throw new NotSupportedException("A string can only be of property type VT_LPWSTR, VT_LPSTR or VT_BSTR.");
            }
        }

        private static long ToPositiveFileTime(DateTime dt)
        {
            var ft = dt.ToUniversalTime().ToFileTimeUtc();
            return ft < 0 ? 0 : ft;
        }

        private bool TryGetVectorValue(VARTYPE vt, out object? value)
        {
            value = null;
            var ret = false;
            int size;
            switch (vt)
            {
                case VARTYPE.VT_LPSTR:
                case VARTYPE.VT_LPWSTR:
                    var strings = new string?[_ca.cElems];
                    for (var i = 0; i < strings.Length; i++)
                    {
                        var str = Marshal.ReadIntPtr(_ca.pElems, IntPtr.Size * i);
                        strings[i] = vt == VARTYPE.VT_LPSTR ? Marshal.PtrToStringAnsi(str) : Marshal.PtrToStringUni(str);
                    }
                    value = strings;
                    ret = true;
                    break;

                case VARTYPE.VT_BOOL:
                    var shorts = new short[_ca.cElems];
                    size = _ca.cElems * Marshal.SizeOf(typeof(short));
                    CopyMemory(Marshal.UnsafeAddrOfPinnedArrayElement(shorts, 0), _ca.pElems, (IntPtr)size);
                    var bools = new bool[shorts.Length];
                    for (var i = 0; i < shorts.Length; i++)
                    {
                        bools[i] = shorts[i] != 0;
                    }
                    value = bools;
                    ret = true;
                    break;

                case VARTYPE.VT_I1:
                case VARTYPE.VT_UI1:
                case VARTYPE.VT_I2:
                case VARTYPE.VT_UI2:
                case VARTYPE.VT_I4:
                case VARTYPE.VT_INT:
                case VARTYPE.VT_UI4:
                case VARTYPE.VT_UINT:
                case VARTYPE.VT_I8:
                case VARTYPE.VT_UI8:
                case VARTYPE.VT_R4:
                case VARTYPE.VT_R8:
                case VARTYPE.VT_ERROR:
                case VARTYPE.VT_CY:
                case VARTYPE.VT_DATE:
                case VARTYPE.VT_FILETIME:
                case VARTYPE.VT_CLSID:
                case VARTYPE.VT_UNKNOWN:
                case VARTYPE.VT_DISPATCH:
                    var et = FromType(vt);
                    var values = Array.CreateInstance(et, _ca.cElems);
                    size = _ca.cElems * Marshal.SizeOf(et);
                    CopyMemory(Marshal.UnsafeAddrOfPinnedArrayElement(values, 0), _ca.pElems, (IntPtr)size);
                    value = values;
                    ret = true;
                    break;
            }
            return ret;
        }

        public IntPtr Serialize(out int size)
        {
            StgSerializePropVariant(this, out var ptr, out size).ThrowOnError();
            return ptr;
        }

        public byte[] Serialize()
        {
            StgSerializePropVariant(this, out var ptr, out var size).ThrowOnError();
            var bytes = new byte[size];
            Marshal.Copy(ptr, bytes, 0, bytes.Length);
            Marshal.FreeCoTaskMem(ptr);
            return bytes;
        }

        public static PROPVARIANT? Deserialize(byte[] bytes) => Deserialize(bytes, true);
        public static PROPVARIANT? Deserialize(byte[] bytes, bool throwOnError)
        {
            if (bytes == null)
                throw new ArgumentNullException(nameof(bytes));

            var pv = new PROPVARIANT();
            var hr = StgDeserializePropVariant(bytes, bytes.Length, pv);
            if (hr.IsError)
            {
                pv.Dispose();
                hr.ThrowOnError(throwOnError);
                return null;
            }

            return pv;
        }

        public static PROPVARIANT? Deserialize(IntPtr ptr, int size) => Deserialize(ptr, size, true);
        public static PROPVARIANT? Deserialize(IntPtr ptr, int size, bool throwOnError)
        {
            if (ptr == IntPtr.Zero)
                throw new ArgumentNullException(nameof(ptr));

            var pv = new PROPVARIANT();
            var hr = StgDeserializePropVariant(ptr, size, pv);
            if (hr.IsError)
            {
                pv.Dispose();
                hr.ThrowOnError(throwOnError);
                return null;
            }

            return pv;
        }

        public override string ToString()
        {
            var value = Value;
            if (value == null)
                return "<null>";

            if (value is string svalue)
                return "[" + VarType + "] `" + svalue + "`";

            if (value is not byte[] && value is IEnumerable enumerable)
                return "[" + VarType + "] " + string.Join(", ", enumerable.OfType<object>());

            if (value is byte[] bytes)
                return "[" + VarType + "] bytes[" + bytes.Length + "]";

            return "[" + VarType + "] " + value;
        }

        [DllImport("propsys", ExactSpelling = true)]
        private extern static HRESULT StgDeserializePropVariant(IntPtr ppProp, int cbMax, [Out] PROPVARIANT ppropvar);

        [DllImport("propsys", ExactSpelling = true)]
        private extern static HRESULT StgDeserializePropVariant([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] ppProp, int cbMax, [Out] PROPVARIANT ppropvar);

        [DllImport("propsys", ExactSpelling = true)]
        private extern static HRESULT StgSerializePropVariant(PROPVARIANT ppropvar, out IntPtr ppProp, out int pcb);

        [DllImport("ole32", ExactSpelling = true)]
        private extern static HRESULT PropVariantClear([In, Out] PROPVARIANT pvar);

        [DllImport("propsys", ExactSpelling = true)]
        private static extern HRESULT InitPropVariantFromFileTime(ref long pftIn, [Out] PROPVARIANT ppropvar);

        [DllImport("kernel32", ExactSpelling = true, EntryPoint = "RtlMoveMemory")]
        internal static extern void CopyMemory(IntPtr destination, IntPtr source, IntPtr length);
    }
}
