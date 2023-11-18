using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using JumpListExplorer.Interop;
using Microsoft.Win32;

namespace JumpListExplorer.Shell
{
    public static class AutomaticDestinationList
    {
        public static IEnumerable<Item> GetItems(string aumid)
        {
            ArgumentNullException.ThrowIfNull(aumid);

            var list = (IAutomaticDestinationList)new CLSID_AutomaticDestinationList();
            var hr = list.Initialize(aumid, null, null);
            if (hr.IsError)
                yield break;

            hr = list.GetList(DESTLISTTYPE.RECENT, int.MaxValue, GETDESTLISTFLAGS.NONE, typeof(IObjectCollection).GUID, out var coll);
            if (hr.IsError)
                yield break;

            coll.GetCount(out var count);
            for (var i = 0; i < count; i++)
            {
                hr = coll.GetAt(i, Native.IID_IUnknown, out var obj);
                if (hr.IsError)
                    continue;

                if (obj is IShellItem item)
                    yield return new Item(item);
            }
        }

        public static int RemoveItems(string aumid, IEnumerable<Item> items)
        {
            ArgumentNullException.ThrowIfNull(aumid);
            ArgumentNullException.ThrowIfNull(items);

            if (!items.Any())
                return 0;

            var list = (IAutomaticDestinationList)new CLSID_AutomaticDestinationList();
            var hr = list.Initialize(aumid, null, null);
            if (hr.IsError)
                return 0;
            hr = list.GetList(DESTLISTTYPE.RECENT, int.MaxValue, GETDESTLISTFLAGS.NONE, typeof(IObjectCollection).GUID, out _);
            if (hr.IsError)
                return 0;

            var count = 0;
            foreach (var item in items)
            {
                hr = list.RemoveDestination(item.NativeObject);
                if (hr.IsSuccess)
                {
                    count++;
                }
            }
            return count;
        }

        public static IEnumerable<string> EnumerateAppUserModelIDs()
        {
            var list = EnumerateAppUserModelIDsFromClassesRoot().ToHashSet();
            foreach (var id in EnumerateAppUserModelIDsFromJumpListData())
            {
                list.Add(id);
            }
            return list;
        }

        private static IEnumerable<string> EnumerateAppUserModelIDsFromJumpListData()
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Search\JumplistData", false);
            if (key == null)
                yield break;

            foreach (var name in key.GetValueNames())
            {
                yield return name;
            }
        }

        private static IEnumerable<string> EnumerateAppUserModelIDsFromClassesRoot()
        {
            using var key = Registry.ClassesRoot;
            if (key == null)
                yield break;

            foreach (var name in key.GetSubKeyNames())
            {
                using var app = key.OpenSubKey(Path.Combine(name, "Application"), false);
                if (app != null)
                {
                    if (app.GetValue("AppUserModelID") is string aumid)
                        yield return aumid;
                }
            }
        }

        [ComImport, Guid("f0ae1542-f497-484b-a175-a20db09144ba")]
        private class CLSID_AutomaticDestinationList { }

        private enum DESTLISTTYPE
        {
            PINNED = 0,
            RECENT = 1,
            FREQUENT = 2
        }

        [Flags]
        private enum GETDESTLISTFLAGS
        {
            NONE = 0x0,
            EXCLUDE_UNNAMED_DESTINATIONS = 0x1,
        }

        [ComImport, Guid("e9c5ef8d-fd41-4f72-ba87-eb03bad5817c"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAutomaticDestinationList
        {
            [PreserveSig]
            HRESULT Initialize([MarshalAs(UnmanagedType.LPWStr)] string appid, [MarshalAs(UnmanagedType.LPWStr)] string? b, [MarshalAs(UnmanagedType.LPWStr)] string? c);

            [PreserveSig]
            HRESULT HasList(out bool has);

            [PreserveSig]
            HRESULT GetList(DESTLISTTYPE type, int count, GETDESTLISTFLAGS flags, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IObjectCollection obj);

            [PreserveSig]
            HRESULT AddUsagePoint([MarshalAs(UnmanagedType.IUnknown)] object obj);

            [PreserveSig]
            HRESULT PinItem([MarshalAs(UnmanagedType.IUnknown)] object obj, bool pinned);

            [PreserveSig]
            HRESULT IsPinned([MarshalAs(UnmanagedType.IUnknown)] object obj, out bool pinned);

            [PreserveSig]
            HRESULT RemoveDestination([MarshalAs(UnmanagedType.IUnknown)] object obj);

            [PreserveSig]
            HRESULT SetUsageData([MarshalAs(UnmanagedType.IUnknown)] object obj, ref float fl, ref long fileTime);

            [PreserveSig]
            HRESULT GetUsageData([MarshalAs(UnmanagedType.IUnknown)] object obj, ref float fl, ref long fileTime);

            [PreserveSig]
            HRESULT ResolveDestination(IntPtr hwnd, int i, [MarshalAs(UnmanagedType.IUnknown)] object shellItem, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object obj);

            [PreserveSig]
            HRESULT ClearList(int i);
        }

        [ComImport, Guid("92ca9dcd-5622-4bba-a805-5e9f541bd8c9"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IObjectArray
        {
            [PreserveSig]
            HRESULT GetCount(out int count);

            [PreserveSig]
            HRESULT GetAt(int index, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object obj);
        }

        [ComImport, Guid("5632b1a4-e38a-400a-928a-d4cd63230295"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IObjectCollection : IObjectArray
        {
            // IObjectArray
            [PreserveSig]
            new HRESULT GetCount(out int count);

            [PreserveSig]
            new HRESULT GetAt(int index, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object obj);

            // IObjectCollection
            [PreserveSig]
            HRESULT AddObject([MarshalAs(UnmanagedType.IUnknown)] object punk);

            [PreserveSig]
            HRESULT AddFromArray(IObjectArray source);

            [PreserveSig]
            HRESULT RemoveObjectAt(int index);

            [PreserveSig]
            HRESULT Clear();
        }
    }
}
