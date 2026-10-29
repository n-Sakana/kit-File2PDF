@echo off
rem Merge PDF and Office files into a PDF for duplex printing.
rem Requires Windows PowerShell 5.1 and the corresponding Office apps.
rem Double-click to choose a folder, or drag a folder onto this file.
setlocal
set "MERGE_SELF=%~f0"
set "MERGE_ARG=%~1"
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:MERGE_SELF,[Text.Encoding]::UTF8); $i=$t.IndexOf('#'+'#PS-BEGIN##',[StringComparison]::Ordinal); Invoke-Expression $t.Substring($i)"
if errorlevel 1 pause
exit /b
##PS-BEGIN##

$BlankAfterLast = $true      # Pad the last document too if it has an odd page count.
$Exts = @('.pdf', '.ppt', '.pptx', '.doc', '.docx', '.xls', '.xlsx', '.xlsm')

$ErrorActionPreference = 'Stop'
$script:ExitCode = 0
$script:Apps = @{}
$script:Warns = New-Object 'System.Collections.Generic.List[string]'
$script:Files = $null
$script:Picked = $null
$script:PickOk = $false
$script:RunOk = $false
$script:PickedDir = $null

$cs = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

// PDF merger compatible with C# 5; no external libraries are required.
// Reads classic xref tables, xref streams and object streams. Writes a classic xref table.
public static class PdfMerge
{
    static readonly Encoding L1 = Encoding.GetEncoding(28591);

    class PdfRef { public int N; public PdfRef(int n) { N = n; } }
    class PdfName { public string V; public PdfName(string v) { V = v; } }
    class PdfStr { public byte[] B; public PdfStr(byte[] b) { B = b; } }
    class PdfNum { public string T; public PdfNum(string t) { T = t; } }
    class PdfNull { public static readonly PdfNull I = new PdfNull(); }
    class PdfDict : Dictionary<string, object> { }
    class PdfStream
    {
        public PdfDict D; public byte[] Raw;
        public PdfStream(PdfDict d, byte[] raw) { D = d; Raw = raw; }
    }
    class PageInfo { public int Old; public PdfDict Node; public PdfDict Inh; }

    static bool IsWs(int c) { return c == 0 || c == 9 || c == 10 || c == 12 || c == 13 || c == 32; }
    static bool IsDelim(int c)
    {
        return c == '(' || c == ')' || c == '<' || c == '>' || c == '[' || c == ']' ||
               c == '{' || c == '}' || c == '/' || c == '%';
    }
    static bool IsDigits(string s)
    {
        if (s.Length == 0) return false;
        for (int i = 0; i < s.Length; i++) { if (s[i] < '0' || s[i] > '9') return false; }
        return true;
    }
    static int HexVal(int c)
    {
        if (c >= '0' && c <= '9') return c - '0';
        if (c >= 'a' && c <= 'f') return c - 'a' + 10;
        if (c >= 'A' && c <= 'F') return c - 'A' + 10;
        return -1;
    }

    class Parser
    {
        public byte[] D; public int P; Doc Owner;
        public Parser(byte[] d, int p, Doc owner) { D = d; P = p; Owner = owner; }

        public void Ws()
        {
            while (P < D.Length)
            {
                int c = D[P];
                if (IsWs(c)) P++;
                else if (c == '%') { while (P < D.Length && D[P] != 10 && D[P] != 13) P++; }
                else break;
            }
        }
        public string Token()
        {
            Ws();
            int st = P;
            while (P < D.Length && !IsWs(D[P]) && !IsDelim(D[P])) P++;
            return L1.GetString(D, st, P - st);
        }
        public object Obj()
        {
            Ws();
            if (P >= D.Length) throw new Exception("Unexpected end of PDF data.");
            int c = D[P];
            if (c == '/') return ReadName();
            if (c == '(') return ReadLitStr();
            if (c == '<')
            {
                if (P + 1 < D.Length && D[P + 1] == '<') return ReadDict();
                return ReadHexStr();
            }
            if (c == '[')
            {
                P++;
                List<object> l = new List<object>();
                while (true)
                {
                    Ws();
                    if (P >= D.Length) throw new Exception("Unterminated PDF array.");
                    if (D[P] == ']') { P++; break; }
                    l.Add(Obj());
                }
                return l;
            }
            string t = Token();
            if (t.Length == 0) throw new Exception("Invalid PDF token at byte offset " + P);
            if (IsDigits(t))
            {
                int save = P;
                Ws();
                string t2 = Token();
                if (IsDigits(t2))
                {
                    Ws();
                    if (P < D.Length && D[P] == 'R' && (P + 1 >= D.Length || IsWs(D[P + 1]) || IsDelim(D[P + 1])))
                    {
                        P++;
                        return new PdfRef(int.Parse(t, CultureInfo.InvariantCulture));
                    }
                }
                P = save;
                return new PdfNum(t);
            }
            if (Regex.IsMatch(t, @"^[+-]?(\d+\.?\d*|\.\d+)$")) return new PdfNum(t);
            if (t == "true") return true;
            if (t == "false") return false;
            if (t == "null") return PdfNull.I;
            throw new Exception("Unexpected PDF token '" + t + "' at byte offset " + P);
        }
        PdfName ReadName()
        {
            P++;
            List<byte> o = new List<byte>();
            while (P < D.Length && !IsWs(D[P]) && !IsDelim(D[P]))
            {
                int c = D[P];
                if (c == '#' && P + 2 < D.Length && HexVal(D[P + 1]) >= 0 && HexVal(D[P + 2]) >= 0)
                {
                    o.Add((byte)(HexVal(D[P + 1]) * 16 + HexVal(D[P + 2])));
                    P += 3;
                    continue;
                }
                o.Add((byte)c); P++;
            }
            return new PdfName(L1.GetString(o.ToArray()));
        }
        PdfStr ReadLitStr()
        {
            P++;
            int depth = 1;
            List<byte> o = new List<byte>();
            while (P < D.Length)
            {
                int c = D[P]; P++;
                if (c == '\\')
                {
                    int e = D[P]; P++;
                    if (e == 'n') o.Add(10);
                    else if (e == 'r') o.Add(13);
                    else if (e == 't') o.Add(9);
                    else if (e == 'b') o.Add(8);
                    else if (e == 'f') o.Add(12);
                    else if (e >= '0' && e <= '7')
                    {
                        int v = e - '0';
                        int k = 0;
                        while (k < 2 && P < D.Length && D[P] >= '0' && D[P] <= '7') { v = v * 8 + (D[P] - '0'); P++; k++; }
                        o.Add((byte)(v & 255));
                    }
                    else if (e == 13) { if (P < D.Length && D[P] == 10) P++; }
                    else if (e == 10) { }
                    else o.Add((byte)e);
                }
                else if (c == '(') { depth++; o.Add((byte)c); }
                else if (c == ')')
                {
                    depth--;
                    if (depth == 0) break;
                    o.Add((byte)c);
                }
                else o.Add((byte)c);
            }
            return new PdfStr(o.ToArray());
        }
        PdfStr ReadHexStr()
        {
            P++;
            List<int> h = new List<int>();
            while (P < D.Length && D[P] != '>')
            {
                int v = HexVal(D[P]);
                if (v >= 0) h.Add(v);
                P++;
            }
            P++;
            if (h.Count % 2 == 1) h.Add(0);
            byte[] b = new byte[h.Count / 2];
            for (int i = 0; i < b.Length; i++) b[i] = (byte)(h[2 * i] * 16 + h[2 * i + 1]);
            return new PdfStr(b);
        }
        PdfDict ReadDict()
        {
            P += 2;
            PdfDict d = new PdfDict();
            while (true)
            {
                Ws();
                if (P + 1 >= D.Length) throw new Exception("Unterminated PDF dictionary.");
                if (D[P] == '>' && D[P + 1] == '>') { P += 2; break; }
                if (D[P] != '/') throw new Exception("Expected a PDF dictionary key at byte offset " + P);
                string k = ReadName().V;
                d[k] = Obj();
            }
            return d;
        }
        // Reads an indirect object starting at the current byte offset.
        public object Indirect(out int num)
        {
            Ws();
            num = int.Parse(Token(), CultureInfo.InvariantCulture);
            Token();
            if (Token() != "obj") throw new Exception("Expected the PDF 'obj' keyword.");
            object o = Obj();
            Ws();
            PdfDict dict = o as PdfDict;
            if (dict != null && P + 6 <= D.Length && L1.GetString(D, P, 6) == "stream")
            {
                P += 6;
                if (D[P] == 13 && D[P + 1] == 10) P += 2;
                else if (D[P] == 10 || D[P] == 13) P++;
                int len = -1;
                if (Owner != null && dict.ContainsKey("Length"))
                {
                    try { len = Owner.ToInt(dict["Length"]); } catch (Exception) { len = -1; }
                }
                bool ok = false;
                if (len >= 0 && P + len <= D.Length)
                {
                    int q = P + len;
                    while (q < D.Length && IsWs(D[q])) q++;
                    ok = (q + 9 <= D.Length && L1.GetString(D, q, 9) == "endstream");
                }
                byte[] raw;
                if (ok)
                {
                    raw = new byte[len];
                    Array.Copy(D, P, raw, 0, len);
                }
                else
                {
                    int e = IndexOf(D, "endstream", P);
                    if (e < 0) throw new Exception("PDF 'endstream' marker not found.");
                    int end = e;
                    if (end - 2 >= P && D[end - 2] == 13 && D[end - 1] == 10) end -= 2;
                    else if (end - 1 >= P && (D[end - 1] == 10 || D[end - 1] == 13)) end -= 1;
                    raw = new byte[end - P];
                    Array.Copy(D, P, raw, 0, end - P);
                }
                return new PdfStream(dict, raw);
            }
            return o;
        }
    }

    static int IndexOf(byte[] d, string s, int from)
    {
        byte[] n = L1.GetBytes(s);
        for (int i = from; i <= d.Length - n.Length; i++)
        {
            int j = 0;
            while (j < n.Length && d[i + j] == n[j]) j++;
            if (j == n.Length) return i;
        }
        return -1;
    }
    static int LastIndexOf(byte[] d, string s)
    {
        byte[] n = L1.GetBytes(s);
        for (int i = d.Length - n.Length; i >= 0; i--)
        {
            int j = 0;
            while (j < n.Length && d[i + j] == n[j]) j++;
            if (j == n.Length) return i;
        }
        return -1;
    }

    class Doc
    {
        public byte[] D;
        public PdfDict Trailer;
        Dictionary<int, int> Off = new Dictionary<int, int>();
        Dictionary<int, int[]> InStm = new Dictionary<int, int[]>();
        Dictionary<int, bool> Seen = new Dictionary<int, bool>();
        Dictionary<int, object> Cache = new Dictionary<int, object>();
        Dictionary<int, bool> StmLoaded = new Dictionary<int, bool>();
        Dictionary<int, bool> Visited = new Dictionary<int, bool>();

        public Doc(byte[] data)
        {
            D = data;
            int i = LastIndexOf(data, "startxref");
            if (i < 0) throw new Exception("PDF 'startxref' marker not found. The file may be damaged or may not be a PDF.");
            Parser p = new Parser(data, i + 9, this);
            ReadXref(int.Parse(p.Token(), CultureInfo.InvariantCulture));
            if (Trailer == null || !Trailer.ContainsKey("Root")) throw new Exception("PDF document root not found.");
            if (Trailer.ContainsKey("Encrypt")) throw new Exception("Encrypted PDFs are not supported.");
        }

        void ReadXref(int off)
        {
            if (Visited.ContainsKey(off)) return;
            Visited[off] = true;
            Parser p = new Parser(D, off, this);
            p.Ws();
            if (p.P + 4 <= D.Length && L1.GetString(D, p.P, 4) == "xref")
            {
                p.P += 4;
                while (true)
                {
                    p.Ws();
                    if (p.P + 7 <= D.Length && L1.GetString(D, p.P, 7) == "trailer") { p.P += 7; break; }
                    int st = int.Parse(p.Token(), CultureInfo.InvariantCulture);
                    int cnt = int.Parse(p.Token(), CultureInfo.InvariantCulture);
                    for (int k = 0; k < cnt; k++)
                    {
                        int o = int.Parse(p.Token(), CultureInfo.InvariantCulture);
                        p.Token();
                        string t = p.Token();
                        int n = st + k;
                        if (Seen.ContainsKey(n)) continue;
                        Seen[n] = true;
                        if (t == "n" && o > 0) Off[n] = o;
                    }
                }
                PdfDict tr = p.Obj() as PdfDict;
                if (tr == null) throw new Exception("Invalid PDF trailer.");
                if (Trailer == null) Trailer = tr;
                if (tr.ContainsKey("XRefStm")) ReadXref(ToInt(tr["XRefStm"]));
                if (tr.ContainsKey("Prev")) ReadXref(ToInt(tr["Prev"]));
            }
            else
            {
                int num;
                PdfStream xs = p.Indirect(out num) as PdfStream;
                if (xs == null) throw new Exception("Expected a PDF cross-reference stream.");
                PdfDict d = xs.D;
                if (Trailer == null) Trailer = d;
                byte[] data = Decode(xs);
                List<object> wl = (List<object>)Res(d["W"]);
                int[] w = new int[wl.Count];
                for (int i = 0; i < w.Length; i++) w[i] = ToInt(wl[i]);
                List<int> idx = new List<int>();
                if (d.ContainsKey("Index"))
                {
                    List<object> il = (List<object>)Res(d["Index"]);
                    for (int i = 0; i < il.Count; i++) idx.Add(ToInt(il[i]));
                }
                else { idx.Add(0); idx.Add(ToInt(d["Size"])); }
                int rl = 0;
                for (int i = 0; i < w.Length; i++) rl += w[i];
                int pos = 0;
                for (int j = 0; j + 1 < idx.Count; j += 2)
                {
                    for (int k = 0; k < idx[j + 1]; k++)
                    {
                        if (pos + rl > data.Length) break;
                        long[] f = new long[3];
                        int q = pos;
                        for (int a = 0; a < 3 && a < w.Length; a++)
                        {
                            long v = 0;
                            for (int b = 0; b < w[a]; b++) { v = (v << 8) | data[q]; q++; }
                            f[a] = v;
                        }
                        pos += rl;
                        long type = (w[0] > 0) ? f[0] : 1;
                        int num2 = idx[j] + k;
                        if (Seen.ContainsKey(num2)) continue;
                        Seen[num2] = true;
                        if (type == 1) Off[num2] = (int)f[1];
                        else if (type == 2) InStm[num2] = new int[] { (int)f[1], (int)f[2] };
                    }
                }
                if (d.ContainsKey("Prev")) ReadXref(ToInt(d["Prev"]));
            }
        }

        public int ToInt(object o)
        {
            o = Res(o);
            PdfNum n = o as PdfNum;
            if (n == null) throw new Exception("Expected a PDF number.");
            return (int)double.Parse(n.T, CultureInfo.InvariantCulture);
        }
        public object Res(object o)
        {
            int guard = 0;
            while (o is PdfRef && guard < 100) { o = Get(((PdfRef)o).N); guard++; }
            return o;
        }
        public object Get(int n)
        {
            object c;
            if (Cache.TryGetValue(n, out c)) return c;
            int off;
            if (Off.TryGetValue(n, out off))
            {
                Parser p = new Parser(D, off, this);
                int num;
                object o = p.Indirect(out num);
                if (num != n) throw new Exception("PDF object number mismatch: " + num + " vs " + n);
                Cache[n] = o;
                return o;
            }
            int[] ins;
            if (InStm.TryGetValue(n, out ins))
            {
                LoadStm(ins[0]);
                if (Cache.TryGetValue(n, out c)) return c;
            }
            return PdfNull.I;
        }
        void LoadStm(int sn)
        {
            if (StmLoaded.ContainsKey(sn)) return;
            StmLoaded[sn] = true;
            PdfStream st = Get(sn) as PdfStream;
            if (st == null) throw new Exception("Expected a PDF object stream.");
            byte[] data = Decode(st);
            int N = ToInt(st.D["N"]);
            int first = ToInt(st.D["First"]);
            Parser p = new Parser(data, 0, null);
            int[] nums = new int[N];
            int[] offs = new int[N];
            for (int i = 0; i < N; i++)
            {
                nums[i] = int.Parse(p.Token(), CultureInfo.InvariantCulture);
                offs[i] = int.Parse(p.Token(), CultureInfo.InvariantCulture);
            }
            for (int i = 0; i < N; i++)
            {
                int[] loc;
                if (InStm.TryGetValue(nums[i], out loc) && loc[0] == sn)
                {
                    Cache[nums[i]] = new Parser(data, first + offs[i], null).Obj();
                }
            }
        }

        public byte[] Decode(PdfStream s)
        {
            object f = s.D.ContainsKey("Filter") ? Res(s.D["Filter"]) : null;
            if (f == null) return s.Raw;
            object dp = s.D.ContainsKey("DecodeParms") ? Res(s.D["DecodeParms"]) : null;
            List<object> fl = f as List<object>;
            if (fl == null) { fl = new List<object>(); fl.Add(f); }
            List<object> dl = dp as List<object>;
            if (dl == null) { dl = new List<object>(); for (int i = 0; i < fl.Count; i++) dl.Add(dp); }
            byte[] data = s.Raw;
            for (int i = 0; i < fl.Count; i++)
            {
                PdfName fn = Res(fl[i]) as PdfName;
                string a = (fn == null) ? "" : fn.V;
                if (a == "FlateDecode" || a == "Fl")
                {
                    data = Inflate(data);
                    PdfDict b = (i < dl.Count) ? (Res(dl[i]) as PdfDict) : null;
                    if (b != null && b.ContainsKey("Predictor") && ToInt(b["Predictor"]) >= 10)
                    {
                        int cols = b.ContainsKey("Columns") ? ToInt(b["Columns"]) : 1;
                        int colors = b.ContainsKey("Colors") ? ToInt(b["Colors"]) : 1;
                        int bpc = b.ContainsKey("BitsPerComponent") ? ToInt(b["BitsPerComponent"]) : 8;
                        data = PngUnpredict(data, cols, colors, bpc);
                    }
                }
                else throw new Exception("Unsupported compression filter in a PDF structure stream: " + a);
            }
            return data;
        }
        static byte[] Inflate(byte[] raw)
        {
            if (raw.Length < 3) return new byte[0];
            MemoryStream outMs = new MemoryStream();
            using (MemoryStream ms = new MemoryStream(raw, 2, raw.Length - 2))
            using (DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress))
            {
                byte[] buf = new byte[8192];
                int n;
                try { while ((n = ds.Read(buf, 0, buf.Length)) > 0) outMs.Write(buf, 0, n); }
                catch (InvalidDataException) { }
            }
            return outMs.ToArray();
        }
        static byte[] PngUnpredict(byte[] data, int cols, int colors, int bpc)
        {
            int bpp = Math.Max(1, colors * bpc / 8);
            int rl = (colors * bpc * cols + 7) / 8;
            int rows = data.Length / (rl + 1);
            byte[] outb = new byte[rows * rl];
            for (int r = 0; r < rows; r++)
            {
                int ft = data[r * (rl + 1)];
                int src = r * (rl + 1) + 1;
                int dst = r * rl;
                for (int i = 0; i < rl; i++)
                {
                    int x = data[src + i];
                    int a = (i >= bpp) ? outb[dst + i - bpp] : 0;
                    int u = (r > 0) ? outb[dst - rl + i] : 0;
                    int c = (r > 0 && i >= bpp) ? outb[dst - rl + i - bpp] : 0;
                    int v;
                    if (ft == 1) v = x + a;
                    else if (ft == 2) v = x + u;
                    else if (ft == 3) v = x + (a + u) / 2;
                    else if (ft == 4)
                    {
                        int pp = a + u - c;
                        int pa = Math.Abs(pp - a), pb = Math.Abs(pp - u), pc = Math.Abs(pp - c);
                        int pr = (pa <= pb && pa <= pc) ? a : ((pb <= pc) ? u : c);
                        v = x + pr;
                    }
                    else v = x;
                    outb[dst + i] = (byte)(v & 255);
                }
            }
            return outb;
        }

        static readonly string[] Inherit = new string[] { "Resources", "MediaBox", "CropBox", "Rotate" };
        public List<PageInfo> Pages()
        {
            List<PageInfo> res = new List<PageInfo>();
            PdfDict root = Res(Trailer["Root"]) as PdfDict;
            if (root == null || !root.ContainsKey("Pages")) throw new Exception("PDF page tree not found.");
            Walk(root["Pages"], new PdfDict(), 0, res, new Dictionary<int, bool>());
            return res;
        }
        void Walk(object r, PdfDict inh, int depth, List<PageInfo> res, Dictionary<int, bool> seen)
        {
            if (depth > 64) throw new Exception("PDF page tree exceeds the depth limit.");
            PdfDict node = Res(r) as PdfDict;
            if (node == null) return;
            PdfDict ni = new PdfDict();
            foreach (KeyValuePair<string, object> kv in inh) ni[kv.Key] = kv.Value;
            for (int i = 0; i < Inherit.Length; i++)
            {
                if (node.ContainsKey(Inherit[i])) ni[Inherit[i]] = node[Inherit[i]];
            }
            if (node.ContainsKey("Kids"))
            {
                PdfRef pr = r as PdfRef;
                if (pr != null)
                {
                    if (seen.ContainsKey(pr.N)) return;
                    seen[pr.N] = true;
                }
                List<object> kids = Res(node["Kids"]) as List<object>;
                if (kids != null)
                {
                    for (int i = 0; i < kids.Count; i++) Walk(kids[i], ni, depth + 1, res, seen);
                }
            }
            else
            {
                PageInfo pi = new PageInfo();
                PdfRef pr2 = r as PdfRef;
                pi.Old = (pr2 != null) ? pr2.N : -1;
                pi.Node = node;
                pi.Inh = ni;
                res.Add(pi);
            }
        }
    }

    class Ctx
    {
        public List<object> Objs; public Doc Src;
        public Dictionary<int, int> Map = new Dictionary<int, int>();
        public Stack<int> Queue = new Stack<int>();
        public int Alloc() { Objs.Add(null); return Objs.Count; }
        public object Cp(object o)
        {
            PdfRef r = o as PdfRef;
            if (r != null)
            {
                int nn;
                if (!Map.TryGetValue(r.N, out nn)) { nn = Alloc(); Map[r.N] = nn; Queue.Push(r.N); }
                return new PdfRef(nn);
            }
            PdfStream st = o as PdfStream;
            if (st != null)
            {
                PdfDict d2 = new PdfDict();
                foreach (KeyValuePair<string, object> kv in st.D) { if (kv.Key != "Length") d2[kv.Key] = kv.Value; }
                return new PdfStream((PdfDict)Cp(d2), st.Raw);
            }
            PdfDict d = o as PdfDict;
            if (d != null)
            {
                PdfDict nd = new PdfDict();
                foreach (KeyValuePair<string, object> kv in d)
                {
                    // The output omits the tag structure tree, so references to it must be removed.
                    if (kv.Key == "StructParent" || kv.Key == "StructParents") continue;
                    nd[kv.Key] = Cp(kv.Value);
                }
                return nd;
            }
            List<object> l = o as List<object>;
            if (l != null)
            {
                List<object> nl = new List<object>();
                for (int i = 0; i < l.Count; i++) nl.Add(Cp(l[i]));
                return nl;
            }
            return o;
        }
        public void Drain()
        {
            while (Queue.Count > 0)
            {
                int q = Queue.Pop();
                object c = Cp(Src.Get(q));
                Objs[Map[q] - 1] = c;
            }
        }
    }

    static void W(Stream w, string s) { byte[] b = L1.GetBytes(s); w.Write(b, 0, b.Length); }
    static string EscName(string n)
    {
        StringBuilder sb = new StringBuilder("/");
        byte[] bs = L1.GetBytes(n);
        for (int i = 0; i < bs.Length; i++)
        {
            int c = bs[i];
            if (c <= 32 || c >= 127 || IsDelim(c) || c == '#') sb.Append('#').Append(c.ToString("X2"));
            else sb.Append((char)c);
        }
        return sb.ToString();
    }
    static void Ser(object o, Stream w)
    {
        if (o is PdfRef) W(w, ((PdfRef)o).N.ToString(CultureInfo.InvariantCulture) + " 0 R");
        else if (o is PdfName) W(w, EscName(((PdfName)o).V));
        else if (o is PdfStr) W(w, "<" + BitConverter.ToString(((PdfStr)o).B).Replace("-", "") + ">");
        else if (o is PdfNum) W(w, ((PdfNum)o).T);
        else if (o is bool) W(w, ((bool)o) ? "true" : "false");
        else if (o is PdfNull) W(w, "null");
        else if (o is PdfStream)
        {
            PdfStream s = (PdfStream)o;
            PdfDict d = new PdfDict();
            foreach (KeyValuePair<string, object> kv in s.D) d[kv.Key] = kv.Value;
            d["Length"] = new PdfNum(s.Raw.Length.ToString(CultureInfo.InvariantCulture));
            Ser(d, w);
            W(w, "\nstream\n");
            w.Write(s.Raw, 0, s.Raw.Length);
            W(w, "\nendstream");
        }
        else if (o is PdfDict)
        {
            W(w, "<<");
            foreach (KeyValuePair<string, object> kv in (PdfDict)o)
            {
                W(w, EscName(kv.Key)); W(w, " "); Ser(kv.Value, w); W(w, " ");
            }
            W(w, ">>");
        }
        else if (o is List<object>)
        {
            W(w, "[");
            foreach (object x in (List<object>)o) { Ser(x, w); W(w, " "); }
            W(w, "]");
        }
        else throw new Exception("Cannot write an unsupported PDF object.");
    }

    static PdfName Nm(string s) { return new PdfName(s); }
    static PdfNum Nn(int i) { return new PdfNum(i.ToString(CultureInfo.InvariantCulture)); }

    const double A4W = 595.276;
    const double A4H = 841.89;

    static string F(double v) { return v.ToString("0.####", CultureInfo.InvariantCulture); }

    static double[] GetBox(Doc doc, PdfDict d, string key)
    {
        if (!d.ContainsKey(key)) return null;
        List<object> l = doc.Res(d[key]) as List<object>;
        if (l == null || l.Count < 4) return null;
        double[] r = new double[4];
        for (int i = 0; i < 4; i++)
        {
            PdfNum n = doc.Res(l[i]) as PdfNum;
            if (n == null) return null;
            r[i] = double.Parse(n.T, CultureInfo.InvariantCulture);
        }
        return r;
    }

    // Returns the visible page size after rotation, or false if the size is unknown.
    static bool VisualSize(Doc doc, PdfDict d, out double vw, out double vh)
    {
        vw = 0; vh = 0;
        double[] box = GetBox(doc, d, "CropBox");
        if (box == null) box = GetBox(doc, d, "MediaBox");
        if (box == null) return false;
        double w0 = Math.Abs(box[2] - box[0]);
        double h0 = Math.Abs(box[3] - box[1]);
        if (w0 < 1 || h0 < 1) return false;
        int rot = 0;
        if (d.ContainsKey("Rotate")) { try { rot = doc.ToInt(d["Rotate"]); } catch (Exception) { rot = 0; } }
        rot = ((rot % 360) + 360) % 360;
        bool swap = (rot == 90 || rot == 270);
        vw = swap ? h0 : w0;
        vh = swap ? w0 : h0;
        return true;
    }

    // Requires a page dictionary containing inherited keys from the source document.
    // Updates its boxes and returns content wrappers for A4 scaling.
    // Returns false for pages already within the A4 tolerance or with an unknown size.
    static bool FitA4(Doc doc, PdfDict d, out string prefix, out string suffix)
    {
        prefix = null; suffix = null;
        double[] box = GetBox(doc, d, "CropBox");
        if (box == null) box = GetBox(doc, d, "MediaBox");
        if (box == null) return false;
        double x0 = Math.Min(box[0], box[2]);
        double y0 = Math.Min(box[1], box[3]);
        double w0 = Math.Abs(box[2] - box[0]);
        double h0 = Math.Abs(box[3] - box[1]);
        if (w0 < 1 || h0 < 1) return false;
        int rot = 0;
        if (d.ContainsKey("Rotate")) { try { rot = doc.ToInt(d["Rotate"]); } catch (Exception) { rot = 0; } }
        rot = ((rot % 360) + 360) % 360;
        bool swap = (rot == 90 || rot == 270);
        double vw = swap ? h0 : w0;
        double vh = swap ? w0 : h0;
        bool land = vw > vh;
        double tw = land ? A4H : A4W;
        double th = land ? A4W : A4H;
        if (Math.Abs(vw - tw) <= 2.0 && Math.Abs(vh - th) <= 2.0) return false;   // Allow a two-point tolerance for A4 dimensions.

        double s = Math.Min(tw / vw, th / vh);
        double bw = swap ? th : tw;      // Page boxes use coordinates before rotation.
        double bh = swap ? tw : th;
        double tx = (bw - s * w0) / 2.0 - s * x0;
        double ty = (bh - s * h0) / 2.0 - s * y0;

        prefix = "q " + F(s) + " 0 0 " + F(s) + " " + F(tx) + " " + F(ty) + " cm " +
                 F(x0) + " " + F(y0) + " " + F(w0) + " " + F(h0) + " re W n\n";
        suffix = "\nQ\n";

        List<object> mb = new List<object>();
        mb.Add(new PdfNum("0")); mb.Add(new PdfNum("0")); mb.Add(new PdfNum(F(bw))); mb.Add(new PdfNum(F(bh)));
        d["MediaBox"] = mb;
        d.Remove("CropBox"); d.Remove("TrimBox"); d.Remove("BleedBox"); d.Remove("ArtBox");
        d.Remove("Annots");   // Annotation positions would no longer match the scaled content.
        return true;
    }

    public static int PageCount(string path)
    {
        Doc d = new Doc(File.ReadAllBytes(path));
        return d.Pages().Count;
    }

    // All arrays must have one entry per input document.
    // Fills startPage with one-based output page numbers, or zero for skipped documents.
    // Fills errors with messages for skipped documents, or null for successful ones.
    // Returns the total output page count, including blank pages.
    public static int Merge(string[] paths, string[] titles, bool[] blankAfter, bool unifyA4,
                            bool skipOnError, string outPath, int[] startPage, string[] errors)
    {
        List<object> objs = new List<object>();
        objs.Add(null);
        int pagesNum = objs.Count;
        List<int> allPages = new List<int>();
        List<int> okDocs = new List<int>();
        List<int> okStarts = new List<int>();

        for (int di = 0; di < paths.Length; di++)
        {
            startPage[di] = 0;
            errors[di] = null;
            int mark = objs.Count;
            int pagesBefore = allPages.Count;
            try
            {
                Doc doc = new Doc(File.ReadAllBytes(paths[di]));
                List<PageInfo> pg = doc.Pages();
                if (pg.Count == 0) throw new Exception("The PDF contains no pages.");

                Ctx cx = new Ctx();
                cx.Objs = objs; cx.Src = doc;
                int[] newNums = new int[pg.Count];
                for (int i = 0; i < pg.Count; i++)
                {
                    newNums[i] = cx.Alloc();
                    if (pg[i].Old >= 0) cx.Map[pg[i].Old] = newNums[i];
                }
                for (int i = 0; i < pg.Count; i++)
                {
                    PdfDict d = new PdfDict();
                    foreach (KeyValuePair<string, object> kv in pg[i].Node) { if (kv.Key != "Parent") d[kv.Key] = kv.Value; }
                    foreach (KeyValuePair<string, object> kv in pg[i].Inh) { if (!d.ContainsKey(kv.Key)) d[kv.Key] = kv.Value; }

                    string pre = null, post = null;
                    List<object> contentRefs = null;
                    if (unifyA4)
                    {
                        object oc = d.ContainsKey("Contents") ? d["Contents"] : null;
                        if (FitA4(doc, d, out pre, out post))
                        {
                            List<object> srcList = new List<object>();
                            if (oc != null)
                            {
                                object rc = doc.Res(oc);
                                if (rc is PdfStream) srcList.Add(oc);
                                else if (rc is List<object>) srcList.AddRange((List<object>)rc);
                            }
                            contentRefs = (List<object>)cx.Cp(srcList);
                            d.Remove("Contents");
                        }
                    }

                    PdfDict copied = (PdfDict)cx.Cp(d);
                    copied["Parent"] = new PdfRef(pagesNum);   // Assign after copying to reference the output page tree.
                    if (contentRefs != null)
                    {
                        int preNum = cx.Alloc();
                        objs[preNum - 1] = new PdfStream(new PdfDict(), L1.GetBytes(pre));
                        int postNum = cx.Alloc();
                        objs[postNum - 1] = new PdfStream(new PdfDict(), L1.GetBytes(post));
                        List<object> all = new List<object>();
                        all.Add(new PdfRef(preNum));
                        all.AddRange(contentRefs);
                        all.Add(new PdfRef(postNum));
                        copied["Contents"] = all;
                    }
                    objs[newNums[i] - 1] = copied;
                    cx.Drain();
                }

                int startIdx = allPages.Count;
                allPages.AddRange(newNums);

                if (blankAfter[di])
                {
                    PdfDict inh = pg[pg.Count - 1].Inh;
                    PdfDict bd = new PdfDict();
                    bd["Type"] = Nm("Page");
                    bd["Parent"] = new PdfRef(pagesNum);
                    bd["Resources"] = new PdfDict();
                    if (unifyA4)
                    {
                        double vw, vh;
                        bool land = VisualSize(doc, inh, out vw, out vh) && vw > vh;
                        List<object> mb = new List<object>();
                        mb.Add(Nn(0)); mb.Add(Nn(0));
                        mb.Add(new PdfNum(F(land ? A4H : A4W)));
                        mb.Add(new PdfNum(F(land ? A4W : A4H)));
                        bd["MediaBox"] = mb;
                    }
                    else
                    {
                        object mbSrc = inh.ContainsKey("MediaBox") ? doc.Res(inh["MediaBox"]) : null;
                        if (mbSrc == null)
                        {
                            List<object> def = new List<object>();
                            def.Add(Nn(0)); def.Add(Nn(0)); def.Add(Nn(612)); def.Add(Nn(792));
                            mbSrc = def;
                        }
                        bd["MediaBox"] = cx.Cp(mbSrc);
                        if (inh.ContainsKey("Rotate")) bd["Rotate"] = cx.Cp(inh["Rotate"]);
                    }
                    int bn = cx.Alloc();
                    objs[bn - 1] = bd;
                    allPages.Add(bn);
                    cx.Drain();
                }

                okDocs.Add(di);
                okStarts.Add(startIdx);
                startPage[di] = startIdx + 1;
            }
            catch (Exception ex)
            {
                // Discard partial copies so a failed document cannot corrupt the output.
                objs.RemoveRange(mark, objs.Count - mark);
                while (allPages.Count > pagesBefore) allPages.RemoveAt(allPages.Count - 1);
                startPage[di] = 0;
                if (!skipOnError) throw new Exception(Path.GetFileName(paths[di]) + ": " + ex.Message);
                errors[di] = ex.Message;
            }
        }

        if (okDocs.Count == 0) throw new Exception("No documents could be merged.");

        PdfDict pagesDict = new PdfDict();
        pagesDict["Type"] = Nm("Pages");
        List<object> kidsOut = new List<object>();
        for (int i = 0; i < allPages.Count; i++) kidsOut.Add(new PdfRef(allPages[i]));
        pagesDict["Kids"] = kidsOut;
        pagesDict["Count"] = Nn(allPages.Count);
        objs[pagesNum - 1] = pagesDict;

        int[] items = new int[okDocs.Count];
        for (int i = 0; i < items.Length; i++) { objs.Add(null); items[i] = objs.Count; }
        objs.Add(null); int ol = objs.Count;
        objs.Add(null); int cat = objs.Count;
        for (int i = 0; i < items.Length; i++)
        {
            PdfDict bm = new PdfDict();
            byte[] tb = Encoding.BigEndianUnicode.GetBytes(titles[okDocs[i]]);
            byte[] withBom = new byte[tb.Length + 2];
            withBom[0] = 0xFE; withBom[1] = 0xFF;
            Array.Copy(tb, 0, withBom, 2, tb.Length);
            bm["Title"] = new PdfStr(withBom);
            bm["Parent"] = new PdfRef(ol);
            List<object> dest = new List<object>();
            dest.Add(new PdfRef(allPages[okStarts[i]])); dest.Add(Nm("Fit"));
            bm["Dest"] = dest;
            if (i > 0) bm["Prev"] = new PdfRef(items[i - 1]);
            if (i < items.Length - 1) bm["Next"] = new PdfRef(items[i + 1]);
            objs[items[i] - 1] = bm;
        }
        PdfDict outl = new PdfDict();
        outl["Type"] = Nm("Outlines");
        outl["First"] = new PdfRef(items[0]);
        outl["Last"] = new PdfRef(items[items.Length - 1]);
        outl["Count"] = Nn(items.Length);
        objs[ol - 1] = outl;
        PdfDict catalog = new PdfDict();
        catalog["Type"] = Nm("Catalog");
        catalog["Pages"] = new PdfRef(pagesNum);
        catalog["Outlines"] = new PdfRef(ol);
        catalog["PageMode"] = Nm("UseOutlines");
        objs[cat - 1] = catalog;

        // Finish writing a temporary PDF before replacing the destination.
        long[] offs = new long[objs.Count];
        string tmpOut = outPath + ".tmp";
        using (FileStream fs = new FileStream(tmpOut, FileMode.Create, FileAccess.Write))
        {
            W(fs, "%PDF-1.7\n");
            fs.Write(new byte[] { 0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A }, 0, 6);
            for (int i = 0; i < objs.Count; i++)
            {
                if (objs[i] == null) throw new Exception("Internal error: PDF object " + (i + 1) + " is empty.");
                offs[i] = fs.Position;
                W(fs, (i + 1).ToString(CultureInfo.InvariantCulture) + " 0 obj\n");
                Ser(objs[i], fs);
                W(fs, "\nendobj\n");
            }
            long xr = fs.Position;
            W(fs, "xref\n0 " + (objs.Count + 1).ToString(CultureInfo.InvariantCulture) + "\n");
            W(fs, "0000000000 65535 f \n");
            for (int i = 0; i < objs.Count; i++)
                W(fs, offs[i].ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n");
            W(fs, "trailer\n<< /Size " + (objs.Count + 1).ToString(CultureInfo.InvariantCulture) +
                  " /Root " + cat.ToString(CultureInfo.InvariantCulture) + " 0 R >>\nstartxref\n" +
                  xr.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        }
        if (File.Exists(outPath)) File.Delete(outPath);
        File.Move(tmpOut, outPath);
        return allPages.Count;
    }
}

public static class TuiUtil
{
    public static int CharWidth(char c)
    {
        int u = (int)c;
        if (char.IsLowSurrogate(c)) return 0;
        if (char.IsHighSurrogate(c)) return 2;
        if ((u >= 0x1100 && u <= 0x115F) || (u >= 0x2E80 && u <= 0xA4CF) || (u >= 0xAC00 && u <= 0xD7A3) ||
            (u >= 0xF900 && u <= 0xFAFF) || (u >= 0xFE30 && u <= 0xFE6F) || (u >= 0xFF00 && u <= 0xFF60) ||
            (u >= 0xFFE0 && u <= 0xFFE6) ||
            // Treat ambiguous symbols as wide to match Japanese console rendering.
            (u >= 0x2190 && u <= 0x21FF) || (u >= 0x2460 && u <= 0x24FF) ||
            (u >= 0x25A0 && u <= 0x25FF) || (u >= 0x2600 && u <= 0x27BF))
            return 2;
        return 1;
    }
    public static int Width(string s)
    {
        if (s == null) return 0;
        int w = 0;
        for (int i = 0; i < s.Length; i++) w += CharWidth(s[i]);
        return w;
    }
    // Truncates the end to fit the requested display width, reserving three columns for dots.
    public static string CutEnd(string s, int maxW)
    {
        if (s == null) return "";
        if (Width(s) <= maxW) return s;
        int keep = maxW - 3;
        if (keep < 0) keep = 0;
        StringBuilder sb = new StringBuilder();
        int w = 0;
        for (int i = 0; i < s.Length; i++)
        {
            int cw = CharWidth(s[i]);
            if (w + cw > keep) break;
            sb.Append(s[i]); w += cw;
        }
        return sb.ToString() + "...";
    }
    // Keeps the filename end of a path when truncating it.
    public static string CutHead(string s, int maxW)
    {
        if (s == null) return "";
        if (Width(s) <= maxW) return s;
        int keep = maxW - 3;
        if (keep < 0) keep = 0;
        StringBuilder sb = new StringBuilder();
        int w = 0;
        for (int i = s.Length - 1; i >= 0; i--)
        {
            int cw = CharWidth(s[i]);
            if (w + cw > keep) break;
            sb.Insert(0, s[i]); w += cw;
        }
        return "..." + sb.ToString();
    }
    // Truncates and pads text to the requested display width.
    public static string Pad(string s, int w)
    {
        string t = CutEnd(s, w);
        int cur = Width(t);
        if (cur < w) t = t + new string(' ', w - cur);
        return t;
    }
}
'@

$script:Self = $env:MERGE_SELF
$script:SelfDir = Split-Path -Parent $script:Self
$script:SettingsPath = [System.IO.Path]::ChangeExtension($script:Self, '.ini')

$script:SettingDefs = @(
    [pscustomobject]@{ Key = 'UnifyA4'; Default = $true; Label = 'Fit pages to A4';
        Desc = 'Resized pages lose links and annotations. Keeps orientation and proportions; blank pages use A4.' },
    [pscustomobject]@{ Key = 'ExcelFitWidth'; Default = $true; Label = 'Fit Excel sheets to one page wide';
        Desc = 'Fits each worksheet to one page wide. Source files are unchanged.' },
    [pscustomobject]@{ Key = 'SkipOnError'; Default = $false; Label = 'Skip files with errors';
        Desc = 'Off: stop on error. On: skip the failed file and continue.' },
    [pscustomobject]@{ Key = 'OpenFolderAfter'; Default = $true; Label = 'Open the output folder when finished';
        Desc = 'Opens File Explorer with the output PDF selected.' },
    [pscustomobject]@{ Key = 'ReturnToMenu'; Default = $false; Label = 'Return to the menu when finished';
        Desc = 'Off: exit when finished. On: return to the menu.' }
)
$script:Cfg = @{}

function Save-Settings {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('# File2PDF settings')
    [void]$sb.AppendLine('# Changes in the Settings menu are saved here automatically.')
    [void]$sb.AppendLine('# Use true or false when editing values manually.')
    [void]$sb.AppendLine('')
    foreach ($d in $script:SettingDefs) {
        [void]$sb.AppendLine('# ' + $d.Label + ' - ' + $d.Desc)
        $v = 'false'
        if ($script:Cfg[$d.Key]) { $v = 'true' }
        [void]$sb.AppendLine($d.Key + '=' + $v)
    }
    try {
        [System.IO.File]::WriteAllText($script:SettingsPath, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
        return $true
    }
    catch { return $false }
}

function Load-Settings {
    foreach ($d in $script:SettingDefs) { $script:Cfg[$d.Key] = $d.Default }
    if (Test-Path -LiteralPath $script:SettingsPath) {
        try {
            foreach ($line in [System.IO.File]::ReadAllLines($script:SettingsPath, [System.Text.Encoding]::UTF8)) {
                $t = $line.Trim()
                if ($t.Length -eq 0 -or $t.StartsWith('#') -or $t.StartsWith(';')) { continue }
                $p = $t.IndexOf('=')
                if ($p -lt 1) { continue }
                $k = $t.Substring(0, $p).Trim()
                $v = $t.Substring($p + 1).Trim().ToLowerInvariant()
                if ($script:Cfg.ContainsKey($k)) {
                    if ($v -eq 'true' -or $v -eq '1' -or $v -eq 'on' -or $v -eq 'yes') { $script:Cfg[$k] = $true }
                    elseif ($v -eq 'false' -or $v -eq '0' -or $v -eq 'off' -or $v -eq 'no') { $script:Cfg[$k] = $false }
                }
            }
        }
        catch { }
    }
    else {
        [void](Save-Settings)
    }
}

function Get-Msg($e) {
    $x = $e.Exception
    while ($x.InnerException) { $x = $x.InnerException }
    return $x.Message
}

function Get-Cols { return [Math]::Max(40, [Console]::WindowWidth - 1) }
function Get-Rows { return [Math]::Max(12, [Console]::WindowHeight - 1) }

function Clear-KeyBuffer {
    while ([Console]::KeyAvailable) { [void][Console]::ReadKey($true) }
}

function mkLine([string]$t, $f = $null, $b = $null) {
    if ($null -eq $f) { $f = $script:DefFg }
    if ($null -eq $b) { $b = $script:DefBg }
    return [pscustomobject]@{ T = $t; F = $f; B = $b }
}

function Draw-Lines($lines) {
    try {
        $cols = Get-Cols
        $rows = Get-Rows
        $top = [Console]::WindowTop
        [Console]::CursorVisible = $false
        for ($i = 0; $i -lt $rows; $i++) {
            [Console]::SetCursorPosition(0, $top + $i)
            if ($i -lt $lines.Count) {
                $l = $lines[$i]
                [Console]::ForegroundColor = $l.F
                [Console]::BackgroundColor = $l.B
                [Console]::Write([TuiUtil]::CutEnd($l.T, $cols - 2))
            }
            else {
                [Console]::ForegroundColor = $script:DefFg
                [Console]::BackgroundColor = $script:DefBg
            }
            # Clear from the actual cursor position to account for console character widths.
            $pad = $cols - [Console]::CursorLeft
            if ($pad -gt 0) { [Console]::Write(' ' * $pad) }
        }
        [Console]::ForegroundColor = $script:DefFg
        [Console]::BackgroundColor = $script:DefBg
    }
    catch {
        [Console]::Clear()    # Retry on the next redraw if the console size changed.
    }
}

function Wait-Enter([string]$text) {
    [Console]::CursorVisible = $true
    Write-Host ''
    Write-Host $text
    Clear-KeyBuffer
    while ($true) {
        $k = [Console]::ReadKey($true)
        if ($k.Key.ToString() -eq 'Enter') { break }
    }
}

function Select-Folder {
    $script:PickedDir = $null
    Add-Type -AssemblyName System.Windows.Forms
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Choose a folder containing files to merge. Subfolders are included.'
    $dlg.ShowNewFolderButton = $false
    $dlg.SelectedPath = $script:SelfDir
    try {
        $r = $dlg.ShowDialog($owner)
        if ($r -eq [System.Windows.Forms.DialogResult]::OK) { $script:PickedDir = $dlg.SelectedPath }
    }
    finally {
        $owner.Dispose()
        $dlg.Dispose()
    }
}

function Get-Targets([string]$dir, [string]$outRegex) {
    $list = New-Object 'System.Collections.Generic.List[object]'
    $root = $dir.TrimEnd('\')
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Recurse -ErrorAction SilentlyContinue)) {
        $ext = $f.Extension.ToLowerInvariant()
        if ($script:Exts -notcontains $ext) { continue }
        if ($f.Name.StartsWith('~$')) { continue }
        $rel = $f.FullName.Substring($root.Length).TrimStart('\')
        if ($rel -match $outRegex) { continue }     # Exclude previous outputs in the selected folder.
        $list.Add([pscustomobject]@{ Rel = $rel; Full = $f.FullName; Ext = $ext; Name = $f.Name })
    }
    $cmp = [System.Comparison[object]] { param($a, $b) [string]::Compare($a.Rel, $b.Rel, [System.StringComparison]::OrdinalIgnoreCase) }
    $list.Sort($cmp)
    $script:Files = $list
}

function Show-ListMenu($title, $infoLines, $listHeader, $listItems, $menuItems, $escText = 'Back') {
    $sel = 0
    $scroll = 0
    [Console]::Clear()
    Clear-KeyBuffer
    while ($true) {
        $cols = Get-Cols
        $rows = Get-Rows
        $n = $listItems.Count
        $mc = $menuItems.Count
        $fixed = 1 + $infoLines.Count + 1 + 1 + $mc + 2
        $listRows = [Math]::Max(1, $rows - $fixed)
        $maxScroll = [Math]::Max(0, $n - $listRows)
        if ($scroll -gt $maxScroll) { $scroll = $maxScroll }
        if ($scroll -lt 0) { $scroll = 0 }

        $lines = New-Object 'System.Collections.Generic.List[object]'
        $lines.Add((mkLine $title 'Yellow'))
        foreach ($t in $infoLines) { $lines.Add((mkLine $t)) }
        $range = ''
        if ($n -gt $listRows) {
            $range = ('   [{0}-{1} / {2}]  PageUp/PageDown: scroll' -f ($scroll + 1), ([Math]::Min($n, $scroll + $listRows)), $n)
        }
        $lines.Add((mkLine ($listHeader + $range) 'Cyan'))
        $last = [Math]::Min($n, $scroll + $listRows)
        for ($i = $scroll; $i -lt $last; $i++) {
            $lines.Add((mkLine ('  ' + $listItems[$i])))
        }
        $lines.Add((mkLine ('-' * $cols) 'DarkGray'))
        for ($i = 0; $i -lt $mc; $i++) {
            if ($i -eq $sel) { $lines.Add((mkLine ('  > ' + $menuItems[$i]) 'Black' 'Gray')) }
            else { $lines.Add((mkLine ('    ' + $menuItems[$i]))) }
        }
        $lines.Add((mkLine ''))
        $lines.Add((mkLine ('  Up/Down: move    Enter: select    Esc: ' + $escText) 'DarkGray'))
        Draw-Lines $lines

        $k = [Console]::ReadKey($true)
        $name = $k.Key.ToString()
        if ($name -eq 'UpArrow') { $sel = ($sel - 1 + $mc) % $mc }
        elseif ($name -eq 'DownArrow') { $sel = ($sel + 1) % $mc }
        elseif ($name -eq 'PageUp') { $scroll = [Math]::Max(0, $scroll - $listRows) }
        elseif ($name -eq 'PageDown') { $scroll = [Math]::Min($maxScroll, $scroll + $listRows) }
        elseif ($name -eq 'Enter') { return $sel }
        elseif ($name -eq 'Escape') { return -1 }
    }
}

function Show-OrderSelect($files) {
    $script:PickOk = $false
    $script:Picked = $null
    $n = $files.Count
    $total = $n + 3
    $order = New-Object 'System.Collections.Generic.List[int]'
    $cur = 0
    $scroll = 0
    $msg = ''
    [Console]::Clear()
    Clear-KeyBuffer
    while ($true) {
        $cols = Get-Cols
        $rows = Get-Rows
        $listRows = [Math]::Max(3, $rows - 6)
        if ($cur -lt $scroll) { $scroll = $cur }
        if ($cur -ge $scroll + $listRows) { $scroll = $cur - $listRows + 1 }
        $maxScroll = [Math]::Max(0, $total - $listRows)
        if ($scroll -gt $maxScroll) { $scroll = $maxScroll }
        if ($scroll -lt 0) { $scroll = 0 }

        $lines = New-Object 'System.Collections.Generic.List[object]'
        $lines.Add((mkLine 'Choose merge order' 'Yellow'))
        $lines.Add((mkLine '  Enter adds/removes files in merge order. Unselected files are excluded.'))
        $lines.Add((mkLine ('-' * $cols) 'DarkGray'))
        $last = [Math]::Min($total, $scroll + $listRows)
        for ($i = $scroll; $i -lt $last; $i++) {
            $fg = $script:DefFg
            if ($i -lt $n) {
                $pos = $order.IndexOf($i)
                if ($pos -ge 0) {
                    $text = ('[{0,2}] ' -f ($pos + 1)) + $files[$i].Rel
                    $fg = 'Green'
                }
                else {
                    $text = '[  ] ' + $files[$i].Rel
                }
            }
            elseif ($i -eq $n) {
                $text = '=== Continue (' + $order.Count + ' selected)'
                $fg = 'Cyan'
            }
            elseif ($i -eq $n + 1) {
                $text = '=== Clear selection'
                $fg = 'Cyan'
            }
            else {
                $text = '=== Back'
                $fg = 'Cyan'
            }
            if ($i -eq $cur) { $lines.Add((mkLine ('> ' + $text) 'Black' 'Gray')) }
            else { $lines.Add((mkLine ('  ' + $text) $fg)) }
        }
        $lines.Add((mkLine ('-' * $cols) 'DarkGray'))
        $lines.Add((mkLine ('  ' + $msg) 'Yellow'))
        $lines.Add((mkLine '  Up/Down: move    Enter: toggle    Home/End/PageUp/PageDown: jump    Esc: back' 'DarkGray'))
        Draw-Lines $lines

        $k = [Console]::ReadKey($true)
        $name = $k.Key.ToString()
        if ($name -eq 'UpArrow') { $cur = ($cur - 1 + $total) % $total }
        elseif ($name -eq 'DownArrow') { $cur = ($cur + 1) % $total }
        elseif ($name -eq 'Home') { $cur = 0 }
        elseif ($name -eq 'End') { $cur = $total - 1 }
        elseif ($name -eq 'PageUp') { $cur = [Math]::Max(0, $cur - $listRows) }
        elseif ($name -eq 'PageDown') { $cur = [Math]::Min($total - 1, $cur + $listRows) }
        elseif ($name -eq 'Escape') { return }
        elseif ($name -eq 'Enter' -or $name -eq 'Spacebar') {
            $msg = ''
            if ($cur -lt $n) {
                $pos = $order.IndexOf($cur)
                if ($pos -ge 0) { $order.RemoveAt($pos) } else { $order.Add($cur) }
            }
            elseif ($cur -eq $n) {
                if ($order.Count -eq 0) { $msg = 'Select at least one file before continuing.' }
                else {
                    $picked = New-Object 'System.Collections.Generic.List[object]'
                    foreach ($ix in $order) { $picked.Add($files[$ix]) }
                    $script:Picked = $picked
                    $script:PickOk = $true
                    return
                }
            }
            elseif ($cur -eq $n + 1) { $order.Clear() }
            else { return }
        }
    }
}

function Show-Settings {
    $defs = $script:SettingDefs
    $cnt = $defs.Count
    $m = $cnt + 1
    $cur = 0
    $msg = ''
    [Console]::Clear()
    Clear-KeyBuffer
    while ($true) {
        $cols = Get-Cols
        $lines = New-Object 'System.Collections.Generic.List[object]'
        $lines.Add((mkLine 'Settings' 'Yellow'))
        $lines.Add((mkLine ('  Settings file: ' + $script:SettingsPath) 'DarkGray'))
        $lines.Add((mkLine ('-' * $cols) 'DarkGray'))
        for ($i = 0; $i -lt $cnt; $i++) {
            $mark = '[ ]'
            if ($script:Cfg[$defs[$i].Key]) { $mark = '[x]' }
            $text = $mark + ' ' + $defs[$i].Label
            if ($i -eq $cur) { $lines.Add((mkLine ('> ' + $text) 'Black' 'Gray')) }
            else { $lines.Add((mkLine ('  ' + $text))) }
        }
        if ($cur -eq $cnt) { $lines.Add((mkLine '> Back' 'Black' 'Gray')) }
        else { $lines.Add((mkLine '  Back')) }
        $lines.Add((mkLine ('-' * $cols) 'DarkGray'))
        if ($cur -lt $cnt) { $lines.Add((mkLine ('  ' + $defs[$cur].Desc) 'Cyan')) }
        else { $lines.Add((mkLine '')) }
        $lines.Add((mkLine ('  ' + $msg) 'Yellow'))
        $lines.Add((mkLine '  Up/Down: move    Enter: toggle and save    Esc: back' 'DarkGray'))
        Draw-Lines $lines

        $k = [Console]::ReadKey($true)
        $name = $k.Key.ToString()
        if ($name -eq 'UpArrow') { $cur = ($cur - 1 + $m) % $m }
        elseif ($name -eq 'DownArrow') { $cur = ($cur + 1) % $m }
        elseif ($name -eq 'Escape') { return }
        elseif ($name -eq 'Enter' -or $name -eq 'Spacebar') {
            if ($cur -ge $cnt) { return }
            $key = $defs[$cur].Key
            $script:Cfg[$key] = -not $script:Cfg[$key]
            if (Save-Settings) { $msg = '' }
            else { $msg = 'Could not save settings. Changes apply only to this session.' }
        }
    }
}

function Convert-Office($f, [string]$outPdf) {
    $ext = $f.Ext
    if ($ext -eq '.ppt' -or $ext -eq '.pptx') {
        if (-not $script:Apps.ppt) {
            try { $script:Apps.ppt = New-Object -ComObject PowerPoint.Application }
            catch { throw 'Could not start PowerPoint. Check that it is installed.' }
        }
        $pres = $script:Apps.ppt.Presentations.Open($f.Full, -1, 0, 0)    # Open read-only without a window.
        try { $pres.SaveAs($outPdf, 32) }                                  # 32 = PDF
        finally { $pres.Close() }
    }
    elseif ($ext -eq '.doc' -or $ext -eq '.docx') {
        if (-not $script:Apps.word) {
            try { $w = New-Object -ComObject Word.Application }
            catch { throw 'Could not start Word. Check that it is installed.' }
            $w.Visible = $false
            try { $w.DisplayAlerts = 0 } catch { }
            try { $w.AutomationSecurity = 3 } catch { }
            $script:Apps.word = $w
        }
        $doc = $script:Apps.word.Documents.Open($f.Full, $false, $true, $false)   # Open read-only without a conversion prompt or recent-file entry.
        try { $doc.ExportAsFixedFormat($outPdf, 17) }                              # 17 = PDF
        finally { $doc.Close(0) }                                                  # 0 = do not save changes.
    }
    elseif ($ext -eq '.xls' -or $ext -eq '.xlsx' -or $ext -eq '.xlsm') {
        if (-not $script:Apps.excel) {
            try { $x = New-Object -ComObject Excel.Application }
            catch { throw 'Could not start Excel. Check that it is installed.' }
            $x.Visible = $false
            $x.DisplayAlerts = $false
            try { $x.EnableEvents = $false } catch { }
            try { $x.AutomationSecurity = 3 } catch { }
            try { $x.AskToUpdateLinks = $false } catch { }
            $script:Apps.excel = $x
        }
        $wb = $script:Apps.excel.Workbooks.Open($f.Full, 0, $true)    # Open read-only without updating links.
        try {
            if ($script:Cfg['ExcelFitWidth']) {
                $failed = 0
                foreach ($ws in $wb.Worksheets) {
                    try {
                        $pgs = $ws.PageSetup
                        $pgs.Zoom = $false
                        $pgs.FitToPagesWide = 1
                        $pgs.FitToPagesTall = $false
                    }
                    catch { $failed++ }
                }
                if ($failed -gt 0) {
                    $script:Warns.Add($f.Rel + ': Could not change print settings for some worksheets, possibly because no printer is configured. Those sheets use their original settings.')
                }
            }
            $wb.ExportAsFixedFormat(0, $outPdf)                         # 0 = PDF
        }
        finally { $wb.Close($false) }
    }
    else {
        throw ('Unsupported file type: ' + $ext)
    }
}

function Close-Apps {
    $a = $script:Apps
    if ($a.ppt) {
        try { if ($a.ppt.Presentations.Count -eq 0) { $a.ppt.Quit() } } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($a.ppt) } catch { }
    }
    if ($a.word) {
        try { if ($a.word.Documents.Count -eq 0) { $a.word.Quit() } } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($a.word) } catch { }
    }
    if ($a.excel) {
        try { if ($a.excel.Workbooks.Count -eq 0) { $a.excel.Quit() } } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($a.excel) } catch { }
    }
    $a.Clear()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Invoke-Merge($sel, [string]$outPath) {
    $script:RunOk = $false
    $script:Warns.Clear()
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('pdfmerge_' + [Guid]::NewGuid().ToString('N'))
    [Console]::Clear()
    [Console]::CursorVisible = $true
    try {
        Write-Host 'Converting and merging...' -ForegroundColor Cyan
        [void](New-Item -ItemType Directory -Path $tmp)
        $skip = [bool]$script:Cfg['SkipOnError']
        $okFiles = New-Object 'System.Collections.Generic.List[object]'
        $okPaths = New-Object 'System.Collections.Generic.List[string]'
        $okCounts = New-Object 'System.Collections.Generic.List[int]'
        $skipped = New-Object 'System.Collections.Generic.List[object]'
        $total = $sel.Count

        for ($i = 0; $i -lt $total; $i++) {
            $f = $sel[$i]
            Write-Host (('[{0}/{1}] ' -f ($i + 1), $total) + $f.Rel)
            $pdf = $null
            $cnt = 0
            $err = $null
            try {
                if ($f.Ext -eq '.pdf') { $pdf = $f.Full }
                else {
                    $pdf = Join-Path $tmp ('c' + $i + '.pdf')
                    Convert-Office $f $pdf
                }
                $cnt = [PdfMerge]::PageCount($pdf)
            }
            catch { $err = Get-Msg $_ }
            if ($null -ne $err) {
                if ($skip) {
                    Write-Host ('    Skipped: ' + $err) -ForegroundColor Yellow
                    $skipped.Add([pscustomobject]@{ Rel = $f.Rel; Reason = $err })
                }
                else { throw ($f.Rel + ' : ' + $err) }
            }
            else {
                $okFiles.Add($f)
                $okPaths.Add($pdf)
                $okCounts.Add($cnt)
            }
        }
        Close-Apps

        if ($okFiles.Count -eq 0) { throw 'No files could be merged.' }
        $n = $okFiles.Count
        $paths = $okPaths.ToArray()
        $titles = New-Object 'string[]' $n
        $blank = New-Object 'bool[]' $n
        for ($i = 0; $i -lt $n; $i++) {
            $t = [System.IO.Path]::GetFileNameWithoutExtension($okFiles[$i].Name)
            $t2 = $t -replace '^\d+[\s_\-　]*', ''
            if ($t2.Length -eq 0) { $t2 = $t }
            $titles[$i] = $t2
            $odd = (($okCounts[$i] % 2) -eq 1)
            $blank[$i] = ($odd -and ($BlankAfterLast -or ($i -lt $n - 1)))
        }
        $start = New-Object 'int[]' $n
        $errs = New-Object 'string[]' $n

        Write-Host ''
        Write-Host 'Merging...'
        $totalPages = [PdfMerge]::Merge($paths, $titles, $blank, [bool]$script:Cfg['UnifyA4'], $skip, $outPath, $start, $errs)

        for ($i = 0; $i -lt $n; $i++) {
            if ($start[$i] -eq 0) {
                $skipped.Add([pscustomobject]@{ Rel = $okFiles[$i].Rel; Reason = $errs[$i] })
            }
        }
        if ($skipped.Count -gt 0) {
            Write-Host ''
            Write-Host ('Skipped files (' + $skipped.Count + '):') -ForegroundColor Yellow
            foreach ($s in $skipped) { Write-Host ('  ' + $s.Rel + ' : ' + $s.Reason) -ForegroundColor Yellow }
        }
        if ($script:Warns.Count -gt 0) {
            Write-Host ''
            Write-Host 'Warnings:' -ForegroundColor Yellow
            foreach ($w in $script:Warns) { Write-Host ('  ' + $w) -ForegroundColor Yellow }
        }
        Write-Host ''
        Write-Host 'Merged documents:'
        for ($i = 0; $i -lt $n; $i++) {
            if ($start[$i] -gt 0) {
                $mark = ''
                if ($blank[$i]) { $mark = ' (+1 blank page)' }
                Write-Host (('  Page {0,3}: {1,3} pages{2}  {3}' -f $start[$i], $okCounts[$i], $mark, $okFiles[$i].Rel))
            }
        }
        Write-Host ''
        Write-Host ('Saved: ' + $outPath + '  (' + $totalPages + ' pages)') -ForegroundColor Green
        $script:RunOk = $true

        if ($script:Cfg['OpenFolderAfter']) {
            try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"' + $outPath + '"') }
            catch { }
        }
    }
    catch {
        Write-Host ''
        Write-Host ('Error: ' + (Get-Msg $_)) -ForegroundColor Red
    }
    finally {
        Close-Apps
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath ($outPath + '.tmp')) { Remove-Item -LiteralPath ($outPath + '.tmp') -Force -ErrorAction SilentlyContinue }
    }
}

function On-Off([bool]$v) { if ($v) { return 'On' } else { return 'Off' } }

function Main {
    $script:DefFg = [Console]::ForegroundColor
    $script:DefBg = [Console]::BackgroundColor

    try { Add-Type -TypeDefinition $cs -Language CSharp }
    catch { throw ('Could not compile the PDF engine: ' + (Get-Msg $_)) }

    Load-Settings

    $dir = $env:MERGE_ARG
    if ([string]::IsNullOrWhiteSpace($dir)) {
        Select-Folder
        $dir = $script:PickedDir
    }
    if ([string]::IsNullOrWhiteSpace($dir)) { return }
    $dir = (Resolve-Path -LiteralPath $dir).ProviderPath

    $trim = $dir.TrimEnd('\')
    $leaf = [System.IO.Path]::GetFileName($trim)
    if ([string]::IsNullOrEmpty($leaf)) { $leaf = $trim.Replace(':', '').Replace('\', '') }
    if ([string]::IsNullOrEmpty($leaf)) { $leaf = 'merged' }
    # Keep this suffix consistent with the exclusion rule for previous outputs.
    $outName = $leaf + '結合版' + (Get-Date -Format 'yyyyMMdd') + '.pdf'
    $outRegex = '^' + [regex]::Escape($leaf) + '結合版\d{8}\.pdf$'
    $outPath = Join-Path $dir $outName

    while ($true) {
        Get-Targets $dir $outRegex
        $files = $script:Files
        if ($files.Count -eq 0) {
            [Console]::Clear()
            Write-Host ('No supported files found (PDF, PowerPoint, Word or Excel): ' + $dir) -ForegroundColor Yellow
            Wait-Enter 'Press Enter to exit.'
            return
        }
        $labels = New-Object 'System.Collections.Generic.List[string]'
        for ($i = 0; $i -lt $files.Count; $i++) { $labels.Add(('{0,3}. {1}' -f ($i + 1), $files[$i].Rel)) }
        $dirShown = [TuiUtil]::CutHead($dir, [Math]::Max(10, (Get-Cols) - 12))
        $info = @(
            ('Folder: ' + $dirShown),
            ('Files: ' + $files.Count + ' (including subfolders, sorted by relative path)')
        )

        $c = Show-ListMenu 'File2PDF' $info 'Files to merge' $labels @('Merge in listed order', 'Choose merge order', 'Settings', 'Exit') 'Exit'
        $sel = $null
        if ($c -eq 0) { $sel = $files }
        elseif ($c -eq 1) {
            Show-OrderSelect $files
            if ($script:PickOk) { $sel = $script:Picked }
        }
        elseif ($c -eq 2) { Show-Settings; continue }
        elseif ($c -eq 3 -or $c -eq -1) { return }
        else { continue }
        if ($null -eq $sel) { continue }

        $exists = Test-Path -LiteralPath $outPath
        $labels2 = New-Object 'System.Collections.Generic.List[string]'
        for ($i = 0; $i -lt $sel.Count; $i++) { $labels2.Add(('{0,3}. {1}' -f ($i + 1), $sel[$i].Rel)) }
        $outShown = [TuiUtil]::CutHead($outPath, [Math]::Max(10, (Get-Cols) - 12))
        $info2 = New-Object 'System.Collections.Generic.List[string]'
        $info2.Add('Output: ' + $outShown)
        $info2.Add('Settings: A4=' + (On-Off $script:Cfg['UnifyA4']) + ' / Excel fit width=' + (On-Off $script:Cfg['ExcelFitWidth']) + ' / Skip errors=' + (On-Off $script:Cfg['SkipOnError']))
        $runLabel = 'Merge'
        if ($exists) {
            $info2.Add('The output file already exists. Merging will overwrite it.')
            $runLabel = 'Overwrite and merge'
        }
        $c2 = Show-ListMenu 'Review merge' $info2 ('Merge order (' + $sel.Count + ' files)') $labels2 @($runLabel, 'Back')
        if ($c2 -ne 0) { continue }

        Invoke-Merge $sel $outPath
        if ($script:Cfg['ReturnToMenu']) { Wait-Enter 'Press Enter to return to the menu.' }
        else {
            Wait-Enter 'Press Enter to exit.'
            return
        }
    }
}

try { Main }
catch {
    [Console]::CursorVisible = $true
    Write-Host ''
    Write-Host ('Error: ' + (Get-Msg $_)) -ForegroundColor Red
    $script:ExitCode = 1
}
finally {
    try { Close-Apps } catch { }
    try { [Console]::CursorVisible = $true; [Console]::ResetColor() } catch { }
}
exit $script:ExitCode
