using System;
using System.Runtime.InteropServices;

namespace JumpListExplorer.Interop
{
    [StructLayout(LayoutKind.Sequential)]
    public struct PROPERTYKEY
    {
        public static readonly PROPERTYKEY Null = new(Guid.Empty, 0);
        public static readonly Guid PSGUID_FOLDER_COLUMNID = new("9e5e05ac-1936-4a75-94f7-4704b8b01923");
        public const int PID_FIRST_USABLE = 2;
        public const int FirstUsableId = PID_FIRST_USABLE;

        public PROPERTYKEY(Guid formatId, int id)
        {
            FormatId = formatId;
            Id = id;
        }

        public Guid FormatId { get; }
        public int Id { get; }
        public readonly bool IsNull => FormatId == Guid.Empty && Id == 0;

        public override readonly string ToString() => FormatId.ToString("B") + " " + Id;

        public static class System
        {
            public static PROPERTYKEY ItemType => new(new Guid("28636aa6-953d-11d2-b5d6-00c04fd918d0"), 11);
            public static PROPERTYKEY ItemTypeText => new(new Guid("b725f130-47ef-101a-a5f1-02608c9eebac"), 4);
            public static PROPERTYKEY Size => new(new Guid("b725f130-47ef-101a-a5f1-02608c9eebac"), 12);
            public static PROPERTYKEY FileAttributes => new(new Guid("b725f130-47ef-101a-a5f1-02608c9eebac"), 13);
            public static PROPERTYKEY DateModified => new(new Guid("b725f130-47ef-101a-a5f1-02608c9eebac"), 14);
            public static PROPERTYKEY DateCreated => new(new Guid("b725f130-47ef-101a-a5f1-02608c9eebac"), 15);
            public static PROPERTYKEY DateAccessed => new(new Guid("b725f130-47ef-101a-a5f1-02608c9eebac"), 16);
        }
    }
}
