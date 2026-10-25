using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using Microsoft.Win32;

namespace AppGroup {
    /// <summary>One entry of an app's jump list: a shell link or a shell item.</summary>
    public class JumpListEntry {
        public string Title { get; set; }
        /// <summary>Absolute PIDL of the target, as returned by ILGetSize/copy.</summary>
        public byte[] TargetIdList { get; set; }
        public string TargetPath { get; set; }
        public string Arguments { get; set; }
        public string WorkingDirectory { get; set; }
        /// <summary>True for a plain shell item (file or folder), false for a shell link.</summary>
        public bool IsShellItem { get; set; }
        internal string Key => $"{TargetPath}|{Arguments}|{(TargetPath == null && TargetIdList != null ? Convert.ToBase64String(TargetIdList) : "")}";
    }

    public class JumpListCategory {
        public string Name { get; set; }
        public bool IsTasks { get; set; }
        public bool IsPinned { get; set; }
        public List<JumpListEntry> Items { get; } = new List<JumpListEntry>();
    }

    /// <summary>
    /// Reads the jump list Windows shows for an app on the taskbar: pinned, recent and
    /// frequent destinations from the shell's automatic destination list, custom categories
    /// and tasks from the app's customDestinations-ms file.
    /// </summary>
    public static class JumpListReader {
        private const uint CategoryFooter = 0xBABFFBAB;

        public static string ResolveAppId(string path) {
            if (string.IsNullOrEmpty(path)) return null;
            try {
                if (path.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase)) {
                    string fromLink = GetLinkAppId(path);
                    if (!string.IsNullOrEmpty(fromLink)) return fromLink;
                    string target = GetLinkTarget(path);
                    if (!string.IsNullOrEmpty(target) && target.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        path = target;
                }
                if (path.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) {
                    // Apps such as VS Code set an explicit AppID at run time; the shell maps the exe to
                    // it through the AppID stored on its Start menu or taskbar shortcut.
                    string fromShortcut = FindShortcutAppIdForExe(path);
                    if (!string.IsNullOrEmpty(fromShortcut)) return fromShortcut;
                }
                return GetDefaultAppId(path);
            }
            catch (Exception ex) {
                Debug.WriteLine($"[JumpList] ResolveAppId failed for {path}: {ex.Message}");
                return null;
            }
        }

        public static List<JumpListCategory> Read(string appId) {
            var result = new List<JumpListCategory>();
            if (string.IsNullOrEmpty(appId)) return result;
            int max = GetMaxItems();

            var pinned = new JumpListCategory { Name = "Pinned", IsPinned = true };
            pinned.Items.AddRange(ReadAutomatic(appId, DESTLISTTYPE.PINNED, int.MaxValue));
            var pinnedKeys = new HashSet<string>(pinned.Items.Select(i => i.Key), StringComparer.OrdinalIgnoreCase);
            if (pinned.Items.Count > 0) result.Add(pinned);

            List<CustomCategory> custom = ReadCustom(appId);
            if (custom == null) {
                // An app with no custom list gets the Recent category by default.
                custom = new List<CustomCategory> { new CustomCategory { KnownId = -2 } };
            }

            JumpListCategory tasks = null;
            foreach (var c in custom) {
                if (c.IsTasks) {
                    tasks = new JumpListCategory { Name = "Tasks", IsTasks = true };
                    tasks.Items.AddRange(c.Items);
                    continue;
                }
                JumpListCategory cat;
                if (c.KnownId != 0) {
                    bool frequent = c.KnownId == -1;
                    cat = new JumpListCategory { Name = frequent ? "Frequent" : "Recent" };
                    cat.Items.AddRange(ReadAutomatic(appId, frequent ? DESTLISTTYPE.FREQUENT : DESTLISTTYPE.RECENT, max + pinnedKeys.Count)
                        .Where(i => !pinnedKeys.Contains(i.Key)).Take(max));
                }
                else {
                    cat = new JumpListCategory { Name = c.Name };
                    cat.Items.AddRange(c.Items.Where(i => !pinnedKeys.Contains(i.Key)).Take(max));
                }
                if (cat.Items.Count > 0) result.Add(cat);
            }
            if (tasks != null && tasks.Items.Count > 0) result.Add(tasks);
            return result;
        }

        /// <summary>Opens an entry the way the taskbar does: links with their arguments, items in their app.</summary>
        public static void Launch(JumpListEntry entry, string ownerPath) {
            if (entry.IsShellItem && !string.IsNullOrEmpty(entry.TargetPath)) {
                string exe = ownerPath;
                if (exe != null && exe.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase)) exe = GetLinkTarget(exe);
                if (exe != null && exe.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) && File.Exists(exe)) {
                    Process.Start(new ProcessStartInfo { FileName = exe, Arguments = $"\"{entry.TargetPath}\"", UseShellExecute = true })?.Dispose();
                    return;
                }
            }

            if (entry.TargetIdList == null) {
                Process.Start(new ProcessStartInfo {
                    FileName = entry.TargetPath,
                    Arguments = entry.Arguments ?? "",
                    WorkingDirectory = entry.WorkingDirectory ?? "",
                    UseShellExecute = true
                })?.Dispose();
                return;
            }

            IntPtr pidl = Marshal.AllocCoTaskMem(entry.TargetIdList.Length);
            try {
                Marshal.Copy(entry.TargetIdList, 0, pidl, entry.TargetIdList.Length);
                var sei = new SHELLEXECUTEINFO {
                    cbSize = Marshal.SizeOf<SHELLEXECUTEINFO>(),
                    fMask = SEE_MASK_IDLIST | SEE_MASK_FLAG_NO_UI,
                    lpIDList = pidl,
                    lpParameters = string.IsNullOrEmpty(entry.Arguments) ? null : entry.Arguments,
                    lpDirectory = string.IsNullOrEmpty(entry.WorkingDirectory) ? null : entry.WorkingDirectory,
                    nShow = 1
                };
                if (!ShellExecuteEx(ref sei))
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            }
            finally {
                Marshal.FreeCoTaskMem(pidl);
            }
        }

        // ---- AppID resolution ----

        private static string GetLinkAppId(string lnkPath) {
            var key = PKEY_AppUserModel_ID;
            Guid iid = typeof(IPropertyStore).GUID;
            if (SHGetPropertyStoreFromParsingName(lnkPath, IntPtr.Zero, GPS_DEFAULT, ref iid, out IPropertyStore store) != 0 || store == null)
                return null;
            try {
                return GetStringProperty(store, key);
            }
            finally {
                Marshal.ReleaseComObject(store);
            }
        }

        private static string GetLinkTarget(string lnkPath) {
            var link = (IShellLinkW)new CShellLink();
            try {
                ((IPersistFile)link).Load(lnkPath, 0);
                var sb = new StringBuilder(1024);
                link.GetPath(sb, sb.Capacity, IntPtr.Zero, 0);
                return sb.ToString();
            }
            catch {
                return null;
            }
            finally {
                Marshal.ReleaseComObject(link);
            }
        }

        private static readonly Dictionary<string, string> _shortcutAppIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static DateTime _shortcutAppIdsTime;

        private static string FindShortcutAppIdForExe(string exePath) {
            lock (_shortcutAppIds) {
                if (DateTime.UtcNow - _shortcutAppIdsTime > TimeSpan.FromMinutes(5)) {
                    _shortcutAppIds.Clear();
                    var folders = new[] {
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"),
                        Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
                        Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu)
                    };
                    foreach (var folder in folders) {
                        if (!Directory.Exists(folder)) continue;
                        IEnumerable<string> links;
                        try { links = Directory.EnumerateFiles(folder, "*.lnk", SearchOption.AllDirectories).ToList(); }
                        catch { continue; }
                        foreach (var lnk in links) {
                            string target = GetLinkTarget(lnk);
                            if (string.IsNullOrEmpty(target) || _shortcutAppIds.ContainsKey(target)) continue;
                            string id = GetLinkAppId(lnk);
                            if (!string.IsNullOrEmpty(id)) _shortcutAppIds[target] = id;
                        }
                    }
                    _shortcutAppIdsTime = DateTime.UtcNow;
                }
                return _shortcutAppIds.TryGetValue(exePath, out var appId) ? appId : null;
            }
        }

        /// <summary>The AppID the shell gives a shortcut to this file when nothing sets one explicitly.</summary>
        private static string GetDefaultAppId(string path) {
            var resolver = (IApplicationResolver)Activator.CreateInstance(Type.GetTypeFromCLSID(CLSID_ApplicationResolver));
            var link = (IShellLinkW)new CShellLink();
            try {
                link.SetPath(path);
                // The resolver dereferences the item, so it must not be null.
                Guid iid = typeof(IShellItem).GUID;
                if (SHCreateItemFromParsingName(path, IntPtr.Zero, ref iid, out IShellItem item) != 0 || item == null) return null;
                int hr = resolver.GetAppIDForShortcutObject(link, item, out IntPtr idPtr);
                Marshal.ReleaseComObject(item);
                if (hr != 0 || idPtr == IntPtr.Zero) return null;
                string id = Marshal.PtrToStringUni(idPtr);
                Marshal.FreeCoTaskMem(idPtr);
                return id;
            }
            finally {
                Marshal.ReleaseComObject(link);
                Marshal.ReleaseComObject(resolver);
            }
        }

        // ---- Automatic destinations ----

        private static List<JumpListEntry> ReadAutomatic(string appId, DESTLISTTYPE type, int maxCount) {
            var items = new List<JumpListEntry>();
            object listObj = null;
            try {
                listObj = Activator.CreateInstance(Type.GetTypeFromCLSID(CLSID_AutomaticDestinationList));
                var list = (IAutomaticDestinationList)listObj;
                list.Initialize(appId, null, null);
                list.HasList(out int hasList);
                if (hasList == 0) return items;
                Guid iid = typeof(IObjectArray).GUID;
                list.GetList(type, maxCount, 0, ref iid, out IObjectArray array);
                if (array != null) {
                    items.AddRange(ReadObjectArray(array));
                    Marshal.ReleaseComObject(array);
                }
            }
            catch (Exception ex) {
                Debug.WriteLine($"[JumpList] automatic {type} for {appId} failed: {ex.Message}");
            }
            finally {
                if (listObj != null) Marshal.ReleaseComObject(listObj);
            }
            return items;
        }

        private static List<JumpListEntry> ReadObjectArray(IObjectArray array) {
            var items = new List<JumpListEntry>();
            array.GetCount(out uint count);
            Guid iidUnknown = new Guid("00000000-0000-0000-C000-000000000046");
            for (uint i = 0; i < count; i++) {
                array.GetAt(i, ref iidUnknown, out object obj);
                if (obj == null) continue;
                try {
                    var entry = ToEntry(obj);
                    if (entry != null) items.Add(entry);
                }
                finally {
                    Marshal.ReleaseComObject(obj);
                }
            }
            return items;
        }

        // ---- Custom destinations ----

        private class CustomCategory {
            public string Name;
            public int KnownId;
            public bool IsTasks;
            public List<JumpListEntry> Items = new List<JumpListEntry>();
        }

        private static List<CustomCategory> ReadCustom(string appId) {
            string file = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                @"Microsoft\Windows\Recent\CustomDestinations", $"{AppIdHash(appId)}.customDestinations-ms");
            if (!File.Exists(file)) return null;
            try {
                byte[] data = File.ReadAllBytes(file);
                using var r = new BinaryReader(new MemoryStream(data), Encoding.Unicode);
                if (r.ReadUInt32() != 2) return null;
                uint categoryCount = r.ReadUInt32();
                r.ReadUInt32();
                var categories = new List<CustomCategory>();
                for (uint c = 0; c < categoryCount; c++) {
                    var cat = new CustomCategory();
                    uint type = r.ReadUInt32();
                    if (type == 0) {
                        int len = r.ReadUInt16();
                        cat.Name = LoadIndirect(new string(r.ReadChars(len)));
                        ReadCustomItems(r, data, cat.Items);
                    }
                    else if (type == 1) {
                        cat.KnownId = r.ReadInt32();
                    }
                    else if (type == 2) {
                        cat.IsTasks = true;
                        ReadCustomItems(r, data, cat.Items);
                    }
                    else {
                        break;
                    }
                    if (r.ReadUInt32() != CategoryFooter) break;
                    categories.Add(cat);
                }
                return categories;
            }
            catch (Exception ex) {
                Debug.WriteLine($"[JumpList] custom destinations for {appId} failed: {ex.Message}");
                return null;
            }
        }

        private static void ReadCustomItems(BinaryReader r, byte[] data, List<JumpListEntry> items) {
            uint count = r.ReadUInt32();
            for (uint i = 0; i < count; i++) {
                var clsid = new Guid(r.ReadBytes(16));
                int start = (int)r.BaseStream.Position;
                object obj = Activator.CreateInstance(Type.GetTypeFromCLSID(clsid));
                IStream stream = SHCreateMemStream(data, (uint)data.Length);
                try {
                    stream.Seek(start, 0, IntPtr.Zero);
                    ((IPersistStream)obj).Load(stream);
                    long pos = 0;
                    unsafe { stream.Seek(0, 1, (IntPtr)(&pos)); }
                    r.BaseStream.Position = pos;
                    var entry = ToEntry(obj);
                    if (entry != null) items.Add(entry);
                }
                finally {
                    Marshal.ReleaseComObject(stream);
                    Marshal.ReleaseComObject(obj);
                }
            }
        }

        // ---- Entries ----

        private static JumpListEntry ToEntry(object obj) {
            if (obj is IShellLinkW link) {
                IPropertyStore store = obj as IPropertyStore;
                if (store != null && GetBoolProperty(store, PKEY_AppUserModel_IsDestListSeparator)) return null;

                var sb = new StringBuilder(2048);
                var entry = new JumpListEntry();
                link.GetPath(sb, sb.Capacity, IntPtr.Zero, 0);
                entry.TargetPath = sb.Length > 0 ? sb.ToString() : null;
                sb.Clear();
                link.GetArguments(sb, sb.Capacity);
                entry.Arguments = sb.ToString();
                sb.Clear();
                link.GetWorkingDirectory(sb, sb.Capacity);
                entry.WorkingDirectory = sb.Length > 0 ? sb.ToString() : null;
                link.GetIDList(out IntPtr pidl);
                if (pidl != IntPtr.Zero) {
                    entry.TargetIdList = CopyIdList(pidl);
                    if (entry.TargetPath == null) entry.TargetPath = GetDisplayName(pidl, SIGDN_DESKTOPABSOLUTEPARSING);
                }

                string title = store != null ? LoadIndirect(GetStringProperty(store, PKEY_Title)) : null;
                if (string.IsNullOrEmpty(title)) {
                    sb.Clear();
                    link.GetDescription(sb, sb.Capacity);
                    title = LoadIndirect(sb.ToString());
                }
                if (string.IsNullOrEmpty(title) && pidl != IntPtr.Zero) title = GetDisplayName(pidl, SIGDN_NORMALDISPLAY);
                if (pidl != IntPtr.Zero) Marshal.FreeCoTaskMem(pidl);
                entry.Title = string.IsNullOrEmpty(title) ? Path.GetFileNameWithoutExtension(entry.TargetPath ?? "") : title;
                return entry;
            }

            if (obj is IShellItem item) {
                if (SHGetIDListFromObject(obj, out IntPtr pidl) != 0 || pidl == IntPtr.Zero) return null;
                try {
                    return new JumpListEntry {
                        IsShellItem = true,
                        TargetIdList = CopyIdList(pidl),
                        TargetPath = GetDisplayName(pidl, SIGDN_FILESYSPATH) ?? GetDisplayName(pidl, SIGDN_DESKTOPABSOLUTEPARSING),
                        Title = GetDisplayName(pidl, SIGDN_NORMALDISPLAY),
                        Arguments = ""
                    };
                }
                finally {
                    Marshal.FreeCoTaskMem(pidl);
                }
            }
            return null;
        }

        private static byte[] CopyIdList(IntPtr pidl) {
            int size = (int)ILGetSize(pidl);
            var bytes = new byte[size];
            Marshal.Copy(pidl, bytes, 0, size);
            return bytes;
        }

        private static string GetDisplayName(IntPtr pidl, uint sigdn) {
            if (SHGetNameFromIDList(pidl, sigdn, out IntPtr name) != 0 || name == IntPtr.Zero) return null;
            string s = Marshal.PtrToStringUni(name);
            Marshal.FreeCoTaskMem(name);
            return s;
        }

        private static string LoadIndirect(string s) {
            if (string.IsNullOrEmpty(s) || s[0] != '@') return s;
            var sb = new StringBuilder(1024);
            return SHLoadIndirectString(s, sb, sb.Capacity, IntPtr.Zero) == 0 ? sb.ToString() : s;
        }

        private static string GetStringProperty(IPropertyStore store, PROPERTYKEY key) {
            var pv = new PROPVARIANT();
            try {
                if (store.GetValue(ref key, out pv) != 0) return null;
                return pv.vt == VT_LPWSTR ? Marshal.PtrToStringUni(pv.p) : null;
            }
            finally {
                PropVariantClear(ref pv);
            }
        }

        private static bool GetBoolProperty(IPropertyStore store, PROPERTYKEY key) {
            var pv = new PROPVARIANT();
            try {
                if (store.GetValue(ref key, out pv) != 0) return false;
                return pv.vt == VT_BOOL && (short)pv.p.ToInt64() != 0;
            }
            finally {
                PropVariantClear(ref pv);
            }
        }

        private static int GetMaxItems() {
            try {
                using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced");
                if (key?.GetValue("JumpListItems_Maximum") is int n && n > 0) return n;
            }
            catch { }
            return 10;
        }

        /// <summary>The shell names jump list files after a CRC-64 of the upper-cased AppID.</summary>
        internal static string AppIdHash(string appId) {
            ulong crc = ulong.MaxValue;
            foreach (byte b in Encoding.Unicode.GetBytes(appId.ToUpperInvariant())) {
                crc ^= b;
                for (int i = 0; i < 8; i++)
                    crc = (crc & 1) != 0 ? (crc >> 1) ^ 0x92C64265D32139A4UL : crc >> 1;
            }
            return crc.ToString("x");
        }

        // ---- Interop ----

        private const uint SEE_MASK_IDLIST = 0x00000004;
        private const uint SEE_MASK_FLAG_NO_UI = 0x00000400;
        private const uint SIGDN_NORMALDISPLAY = 0x00000000;
        private const uint SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000;
        private const uint SIGDN_FILESYSPATH = 0x80058000;
        private const int GPS_DEFAULT = 0;
        private const ushort VT_LPWSTR = 31;
        private const ushort VT_BOOL = 11;

        private static readonly Guid CLSID_AutomaticDestinationList = new Guid("F0AE1542-F497-484B-A175-A20DB09144BA");
        private static readonly Guid CLSID_ApplicationResolver = new Guid("660B90C8-73A9-4B58-8CAE-355B7F55341B");

        private static readonly PROPERTYKEY PKEY_AppUserModel_ID = new PROPERTYKEY(new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), 5);
        private static readonly PROPERTYKEY PKEY_AppUserModel_IsDestListSeparator = new PROPERTYKEY(new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), 6);
        private static readonly PROPERTYKEY PKEY_Title = new PROPERTYKEY(new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9"), 2);

        private enum DESTLISTTYPE { PINNED = 0, RECENT = 1, FREQUENT = 2 }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        private struct PROPERTYKEY {
            public Guid fmtid;
            public uint pid;
            public PROPERTYKEY(Guid f, uint p) { fmtid = f; pid = p; }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct PROPVARIANT {
            public ushort vt;
            public ushort r1, r2, r3;
            public IntPtr p;
            public IntPtr p2;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct SHELLEXECUTEINFO {
            public int cbSize;
            public uint fMask;
            public IntPtr hwnd;
            public string lpVerb;
            public string lpFile;
            public string lpParameters;
            public string lpDirectory;
            public int nShow;
            public IntPtr hInstApp;
            public IntPtr lpIDList;
            public string lpClass;
            public IntPtr hkeyClass;
            public uint dwHotKey;
            public IntPtr hIcon;
            public IntPtr hProcess;
        }

        /// <summary>Undocumented shell interface behind the taskbar's pinned/recent/frequent lists (Windows 10+).</summary>
        [ComImport, Guid("E9C5EF8D-FD41-4F72-BA87-EB03BAD5817C"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAutomaticDestinationList {
            void Initialize([MarshalAs(UnmanagedType.LPWStr)] string appUserModelId, [MarshalAs(UnmanagedType.LPWStr)] string lnkPath, [MarshalAs(UnmanagedType.LPWStr)] string unknown);
            void HasList(out int hasList);
            void GetList(DESTLISTTYPE type, int maxCount, int flags, ref Guid riid, out IObjectArray ppv);
        }

        /// <summary>Undocumented shell interface that computes AppIDs (Windows 8+).</summary>
        [ComImport, Guid("DE25675A-72DE-44B4-9373-05170450C140"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IApplicationResolver {
            [PreserveSig] int GetAppIDForShortcut(IShellItem psi, out IntPtr appId);
            [PreserveSig] int GetAppIDForShortcutObject(IShellLinkW psl, IShellItem psi, out IntPtr appId);
        }

        [ComImport, Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IObjectArray {
            void GetCount(out uint count);
            void GetAt(uint index, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);
        }

        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem { }

        [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPropertyStore {
            void GetCount(out uint count);
            void GetAt(uint index, out PROPERTYKEY key);
            [PreserveSig] int GetValue(ref PROPERTYKEY key, out PROPVARIANT pv);
        }

        [ComImport, Guid("000214F9-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellLinkW {
            void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cch, IntPtr pfd, uint fFlags);
            void GetIDList(out IntPtr ppidl);
            void SetIDList(IntPtr pidl);
            void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);
            void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);
            void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
            void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);
            void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
            void GetHotkey(out short pwHotkey);
            void SetHotkey(short wHotkey);
            void GetShowCmd(out int piShowCmd);
            void SetShowCmd(int iShowCmd);
            void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cch, out int piIcon);
            void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
            void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
            void Resolve(IntPtr hwnd, uint fFlags);
            void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
        }

        [ComImport, Guid("00021401-0000-0000-C000-000000000046")]
        private class CShellLink { }

        [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPersistStream {
            void GetClassID(out Guid pClassID);
            [PreserveSig] int IsDirty();
            void Load(IStream pStm);
            void Save(IStream pStm, [MarshalAs(UnmanagedType.Bool)] bool fClearDirty);
            void GetSizeMax(out long pcbSize);
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHGetPropertyStoreFromParsingName(string pszPath, IntPtr pbc, int flags, ref Guid riid, out IPropertyStore ppv);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHCreateItemFromParsingName(string pszPath, IntPtr pbc, ref Guid riid, out IShellItem ppv);

        [DllImport("shell32.dll")]
        private static extern int SHGetIDListFromObject([MarshalAs(UnmanagedType.IUnknown)] object punk, out IntPtr ppidl);

        [DllImport("shell32.dll")]
        private static extern int SHGetNameFromIDList(IntPtr pidl, uint sigdnName, out IntPtr ppszName);

        [DllImport("shell32.dll")]
        private static extern uint ILGetSize(IntPtr pidl);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool ShellExecuteEx(ref SHELLEXECUTEINFO lpExecInfo);

        [DllImport("shlwapi.dll")]
        private static extern IStream SHCreateMemStream(byte[] pInit, uint cbInit);

        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        private static extern int SHLoadIndirectString(string pszSource, StringBuilder pszOutBuf, int cchOutBuf, IntPtr ppvReserved);

        [DllImport("ole32.dll")]
        private static extern int PropVariantClear(ref PROPVARIANT pvar);
    }
}
