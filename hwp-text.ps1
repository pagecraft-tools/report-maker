# Direct HWP 5.0 (OLE compound file) text extractor - no Hangul COM needed.
$ErrorActionPreference = 'Stop'

$cs = @"
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

[ComImport, Guid("0000000B-0000-0000-C000-000000000046"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IStorage {
  void CreateStream(string n, uint m, uint r1, uint r2, out IStream s);
  void OpenStream(string n, IntPtr r1, uint m, uint r2, out IStream s);
  void CreateStorage(string n, uint m, uint r1, uint r2, out IStorage s);
  void OpenStorage(string n, IStorage pri, uint m, IntPtr snb, uint r, out IStorage s);
  void CopyTo(uint c, Guid[] iid, IntPtr snb, IStorage dest);
  void MoveElementTo(string n, IStorage dest, string newName, uint flags);
  void Commit(uint flags);
  void Revert();
  void EnumElements(uint r1, IntPtr r2, uint r3, out IntPtr e);
  void DestroyElement(string n);
  void RenameElement(string oldName, string newName);
  void SetElementTimes(string n, IntPtr c, IntPtr a, IntPtr m);
  void SetClass(Guid clsid);
  void SetStateBits(uint bits, uint mask);
  void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG s, uint flag);
}

public static class Hwp {
  [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
  static extern int StgOpenStorage(string name, IStorage pri, uint mode, IntPtr snb, uint r,
                                   [MarshalAs(UnmanagedType.Interface)] out IStorage stg);

  const uint MODE = 0x00000020; // STGM_READ(0) | STGM_SHARE_DENY_WRITE(0x20)

  static byte[] ReadStream(IStorage stg, string name) {
    IStream st = null;
    Exception last = null;
    foreach (uint m in new uint[] { 0x20, 0x10, 0x00 }) {
      try { stg.OpenStream(name, IntPtr.Zero, m, 0, out st); last = null; break; }
      catch (Exception ex) { last = ex; st = null; }
    }
    if (st == null) throw new Exception("OpenStream(" + name + ") failed: " + last.Message);
    System.Runtime.InteropServices.ComTypes.STATSTG s;
    st.Stat(out s, 1);
    int len = (int)s.cbSize;
    byte[] buf = new byte[len];
    IntPtr read = Marshal.AllocHGlobal(8);
    st.Read(buf, len, read);
    Marshal.FreeHGlobal(read);
    Marshal.ReleaseComObject(st);
    return buf;
  }

  static List<string> Names(IStorage stg) {
    // Sections are always named Section0, Section1, ... - probe until one is missing.
    var list = new List<string>();
    for (int i = 0; i < 64; i++) {
      string n = "Section" + i;
      try { ReadStream(stg, n); list.Add(n); }
      catch { break; }
    }
    return list;
  }

  static byte[] Inflate(byte[] data) {
    using (var ms = new MemoryStream(data))
    using (var ds = new DeflateStream(ms, CompressionMode.Decompress))
    using (var outMs = new MemoryStream()) {
      byte[] b = new byte[65536]; int n;
      while ((n = ds.Read(b, 0, b.Length)) > 0) outMs.Write(b, 0, n);
      return outMs.ToArray();
    }
  }

  static bool IsWide(int c) {
    // INLINE (4,5,6,7,8,9,19,20) and EXTENDED (1,2,3,11,12,14,15,16,17,18,21,22,23)
    // both occupy 8 wchars. Everything else in 0..31 occupies 1 wchar.
    switch (c) {
      case 1: case 2: case 3: case 4: case 5: case 6: case 7: case 8: case 9:
      case 11: case 12: case 14: case 15: case 16: case 17: case 18:
      case 19: case 20: case 21: case 22: case 23:
        return true;
      default: return false;
    }
  }

  static void ParseSection(byte[] d, List<string> outParas) {
    int p = 0;
    while (p + 4 <= d.Length) {
      uint h = BitConverter.ToUInt32(d, p); p += 4;
      int tag = (int)(h & 0x3FF);
      int size = (int)((h >> 20) & 0xFFF);
      if (size == 0xFFF) { size = (int)BitConverter.ToUInt32(d, p); p += 4; }
      if (size < 0 || p + size > d.Length) break;
      if (tag == 67) {
        var sb = new StringBuilder();
        int i = p, end = p + size;
        while (i + 1 < end) {
          ushort c = BitConverter.ToUInt16(d, i);
          if (c < 32) {
            if (c == 9) { sb.Append('\t'); i += 16; continue; }
            if (IsWide(c)) {
              if (c == 11) sb.Append("{{OBJ}}");
              i += 16; continue;
            }
            if (c == 10) { sb.Append('\n'); i += 2; continue; }
            i += 2; continue;
          }
          sb.Append((char)c);
          i += 2;
        }
        outParas.Add(sb.ToString());
      }
      p += size;
    }
  }

  public static List<string> Trace = new List<string>();

  public static string Extract(string path) {
    Trace.Add("begin");
    IStorage root;
    int hr = StgOpenStorage(path, null, MODE, IntPtr.Zero, 0, out root);
    if (hr != 0) throw new Exception("StgOpenStorage hr=0x" + hr.ToString("X8"));
    Trace.Add("root opened");
    byte[] fh = ReadStream(root, "FileHeader");
    Trace.Add("FileHeader read len=" + fh.Length);
    uint flags = BitConverter.ToUInt32(fh, 36);
    bool compressed = (flags & 1) != 0;

    IStorage body;
    root.OpenStorage("BodyText", null, 0x10, IntPtr.Zero, 0, out body); // STGM_READ|STGM_SHARE_EXCLUSIVE
    Trace.Add("BodyText opened");
    var names = Names(body);
    names.Sort(StringComparer.Ordinal);
    var sb = new StringBuilder();
    sb.AppendLine("### compressed=" + compressed + " sections=" + names.Count);
    foreach (var n in names) {
      byte[] raw = ReadStream(body, n);
      byte[] data = compressed ? Inflate(raw) : raw;
      var list = new List<string>();
      ParseSection(data, list);
      sb.AppendLine("### " + n + " bytes=" + data.Length + " paras=" + list.Count);
      foreach (var t in list) sb.AppendLine(t);
    }
    Marshal.ReleaseComObject(body);
    Marshal.ReleaseComObject(root);
    return sb.ToString();
  }
}
"@

Add-Type -TypeDefinition $cs -Language CSharp
$txt = [Hwp]::Extract($args[0])
[System.IO.File]::WriteAllText($args[1], $txt, (New-Object System.Text.UTF8Encoding $false))
Write-Output ("chars: " + $txt.Length)
