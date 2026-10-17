@echo off
rem ============================================================
rem  File2PDF.cmd
rem  Merge PDF / PowerPoint / Word / Excel files into one PDF
rem  for duplex printing (text UI, no admin rights, no extra install)
rem
rem  Usage : double-click, then pick a folder in the dialog
rem          (or drag a folder onto this file)
rem  Needs : Windows PowerShell 5.1 + Office (only for Office files)
rem  Config: File2PDF.ini is created next to this file
rem ============================================================
setlocal
set "MERGE_SELF=%~f0"
set "MERGE_ARG=%~1"
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:MERGE_SELF,[Text.Encoding]::UTF8); $i=$t.IndexOf('#'+'#PS-BEGIN##',[StringComparison]::Ordinal); Invoke-Expression $t.Substring($i)"
if errorlevel 1 pause
exit /b
##PS-BEGIN##
# ============================================================
#  ここから PowerShell。上の cmd 部分から呼び出されます。
#  操作: 上下キーで移動 / Enter で決定 / Esc で戻る
# ============================================================

# ---------- 固定設定（変えたいときはここだけ編集） ----------
$BlankAfterLast = $true      # 最後の資料が奇数ページでも後ろに白紙を入れる
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

# ---------- C# : PDF 結合エンジン + 画面幅ユーティリティ（外部ライブラリ不要） ----------
$cs = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

// Minimal PDF merger (C# 5 compatible, no external libraries).
// Reads classic xref tables, xref streams and object streams. Writes a classic xref table.
public static class PdfMerge
{
    static readonly Encoding L1 = Encoding.GetEncoding(28591);

    // ---------- object model ----------
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

    // ---------- parser ----------
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
            if (P >= D.Length) throw new Exception("unexpected end of data");
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
                    if (P >= D.Length) throw new Exception("unterminated array");
                    if (D[P] == ']') { P++; break; }
                    l.Add(Obj());
                }
                return l;
            }
            string t = Token();
            if (t.Length == 0) throw new Exception("bad token at " + P);
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
            throw new Exception("unexpected token '" + t + "' at " + P);
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
                if (P + 1 >= D.Length) throw new Exception("unterminated dictionary");
                if (D[P] == '>' && D[P + 1] == '>') { P += 2; break; }
                if (D[P] != '/') throw new Exception("dictionary key expected at " + P);
                string k = ReadName().V;
                d[k] = Obj();
            }
            return d;
        }
        // Parses "N G obj ... endobj" at current position.
        public object Indirect(out int num)
        {
            Ws();
            num = int.Parse(Token(), CultureInfo.InvariantCulture);
            Token();
            if (Token() != "obj") throw new Exception("'obj' expected");
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
                    if (e < 0) throw new Exception("endstream not found");
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

    // ---------- document ----------
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
            if (i < 0) throw new Exception("startxref not found (not a PDF or damaged)");
            Parser p = new Parser(data, i + 9, this);
            ReadXref(int.Parse(p.Token(), CultureInfo.InvariantCulture));
            if (Trailer == null || !Trailer.ContainsKey("Root")) throw new Exception("Root not found");
            if (Trailer.ContainsKey("Encrypt")) throw new Exception("encrypted (password protected) PDF is not supported");
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
                if (tr == null) throw new Exception("bad trailer");
                if (Trailer == null) Trailer = tr;
                if (tr.ContainsKey("XRefStm")) ReadXref(ToInt(tr["XRefStm"]));
                if (tr.ContainsKey("Prev")) ReadXref(ToInt(tr["Prev"]));
            }
            else
            {
                int num;
                PdfStream xs = p.Indirect(out num) as PdfStream;
                if (xs == null) throw new Exception("xref stream expected");
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
            if (n == null) throw new Exception("number expected");
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
                if (num != n) throw new Exception("object number mismatch " + num + " vs " + n);
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
            if (st == null) throw new Exception("object stream expected");
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
                else throw new Exception("unsupported filter in structure stream: " + a);
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
            if (root == null || !root.ContainsKey("Pages")) throw new Exception("page tree not found");
            Walk(root["Pages"], new PdfDict(), 0, res, new Dictionary<int, bool>());
            return res;
        }
        void Walk(object r, PdfDict inh, int depth, List<PageInfo> res, Dictionary<int, bool> seen)
        {
            if (depth > 64) throw new Exception("page tree too deep");
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

    // ---------- copying objects into the output ----------
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
                    // the tag structure tree is not carried over -> drop links to it
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

    // ---------- writer ----------
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
        else throw new Exception("cannot serialize object");
    }

    static PdfName Nm(string s) { return new PdfName(s); }
    static PdfNum Nn(int i) { return new PdfNum(i.ToString(CultureInfo.InvariantCulture)); }

    // ---------- A4 unification ----------
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

    // Visible (rotated) size of a page. Returns false if it cannot be determined.
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

    // Fits one page (dictionary d, in SOURCE object space, already containing inherited keys)
    // onto an A4 page. Returns false when nothing had to be changed (already A4 / unknown size).
    // On true: d is modified (MediaBox etc.), and the new content-stream pieces are returned.
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
        if (Math.Abs(vw - tw) <= 2.0 && Math.Abs(vh - th) <= 2.0) return false;   // already A4

        double s = Math.Min(tw / vw, th / vh);
        double bw = swap ? th : tw;      // page box in unrotated space
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
        d.Remove("Annots");   // link positions would no longer match after scaling
        return true;
    }

    // ---------- public API ----------
    public static int PageCount(string path)
    {
        Doc d = new Doc(File.ReadAllBytes(path));
        return d.Pages().Count;
    }

    // paths/titles/blankAfter : one entry per input document.
    // startPage[i] (out) : 1-based first page of document i in the result, 0 if it was skipped.
    // errors[i]      (out) : error text when document i was skipped, otherwise null.
    // Returns the total number of pages written.
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
                if (pg.Count == 0) throw new Exception("no pages");

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
                    copied["Parent"] = new PdfRef(pagesNum);   // set AFTER copy: must point to the new Pages node
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
                // roll back everything this document added
                objs.RemoveRange(mark, objs.Count - mark);
                while (allPages.Count > pagesBefore) allPages.RemoveAt(allPages.Count - 1);
                startPage[di] = 0;
                if (!skipOnError) throw new Exception(Path.GetFileName(paths[di]) + ": " + ex.Message);
                errors[di] = ex.Message;
            }
        }

        if (okDocs.Count == 0) throw new Exception("no document could be merged");

        PdfDict pagesDict = new PdfDict();
        pagesDict["Type"] = Nm("Pages");
        List<object> kidsOut = new List<object>();
        for (int i = 0; i < allPages.Count; i++) kidsOut.Add(new PdfRef(allPages[i]));
        pagesDict["Kids"] = kidsOut;
        pagesDict["Count"] = Nn(allPages.Count);
        objs[pagesNum - 1] = pagesDict;

        // bookmarks (only for documents that were merged)
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

        // write (to a temp file first, so a failure never leaves a broken PDF behind)
        long[] offs = new long[objs.Count];
        string tmpOut = outPath + ".tmp";
        using (FileStream fs = new FileStream(tmpOut, FileMode.Create, FileAccess.Write))
        {
            W(fs, "%PDF-1.7\n");
            fs.Write(new byte[] { 0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A }, 0, 6);
            for (int i = 0; i < objs.Count; i++)
            {
                if (objs[i] == null) throw new Exception("internal error: object " + (i + 1) + " is empty");
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

// ---------- helpers for the text UI (display width of full-width characters) ----------
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
            // East Asian "ambiguous" symbols (circled numbers, arrows, shapes...): wide in Japanese consoles
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
    // Cut the end so that the display width is <= maxW.
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
    // Keep the tail (file name side) of a path: "...rest\of\path.pdf"
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
    // Cut (end) and pad with spaces to exactly w display columns.
    public static string Pad(string s, int w)
    {
        string t = CutEnd(s, w);
        int cur = Width(t);
        if (cur < w) t = t + new string(' ', w - cur);
        return t;
    }
}
'@

# ---------- 設定ファイル（この cmd と同じフォルダ。なければ自動生成） ----------
$script:Self = $env:MERGE_SELF
$script:SelfDir = Split-Path -Parent $script:Self
$script:SettingsPath = [System.IO.Path]::ChangeExtension($script:Self, '.ini')

$script:SettingDefs = @(
    [pscustomobject]@{ Key = 'UnifyA4'; Default = $true; Label = 'ページサイズを A4 に統一';
        Desc = '縦横の向きと比率を保ったまま A4 に収めます。白紙も A4 になります（PDF 内のリンクは無効になります）' },
    [pscustomobject]@{ Key = 'ExcelFitWidth'; Default = $true; Label = 'Excel を横 1 ページに収める';
        Desc = 'Excel の各シートを横 1 ページに収めて変換します（元のファイルは変更しません）' },
    [pscustomobject]@{ Key = 'SkipOnError'; Default = $false; Label = '変換に失敗したらスキップ';
        Desc = 'オフ: 失敗したら中止してエラーを表示 / オン: そのファイルを除いて続行' },
    [pscustomobject]@{ Key = 'OpenFolderAfter'; Default = $true; Label = '完了後に出力フォルダを開く';
        Desc = '出力した PDF を選択した状態でエクスプローラーを開きます' },
    [pscustomobject]@{ Key = 'ReturnToMenu'; Default = $false; Label = '完了後にメニューへ戻る';
        Desc = 'オフ: 完了したら終了 / オン: メニューに戻って続けて操作できます' }
)
$script:Cfg = @{}

function Save-Settings {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('# File2PDF 設定ファイル')
    [void]$sb.AppendLine('# メニューの「設定」で変更すると、この内容が自動で更新されます。')
    [void]$sb.AppendLine('# 手で編集する場合は true / false で指定してください。')
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
        [void](Save-Settings)    # 初回: 既定値で生成
    }
}

# ---------- 共通 ----------
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
            # 実際のカーソル位置から行末まで空白で埋める（全角・記号の幅の差があっても崩れない）
            $pad = $cols - [Console]::CursorLeft
            if ($pad -gt 0) { [Console]::Write(' ' * $pad) }
        }
        [Console]::ForegroundColor = $script:DefFg
        [Console]::BackgroundColor = $script:DefBg
    }
    catch {
        [Console]::Clear()    # ウィンドウサイズ変更などで失敗したら次の描画でやり直す
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

# ---------- ① フォルダ選択ダイアログ ----------
function Select-Folder {
    $script:PickedDir = $null
    Add-Type -AssemblyName System.Windows.Forms
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'まとめたいファイルが入っているフォルダを選んでください（サブフォルダも対象です）'
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

# ---------- ② ファイル収集（サブフォルダ込み・昇順） ----------
function Get-Targets([string]$dir, [string]$outRegex) {
    $list = New-Object 'System.Collections.Generic.List[object]'
    $root = $dir.TrimEnd('\')
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Recurse -ErrorAction SilentlyContinue)) {
        $ext = $f.Extension.ToLowerInvariant()
        if ($script:Exts -notcontains $ext) { continue }
        if ($f.Name.StartsWith('~$')) { continue }
        $rel = $f.FullName.Substring($root.Length).TrimStart('\')
        if ($rel -match $outRegex) { continue }     # 前回までの出力ファイルは除外
        $list.Add([pscustomobject]@{ Rel = $rel; Full = $f.FullName; Ext = $ext; Name = $f.Name })
    }
    $cmp = [System.Comparison[object]] { param($a, $b) [string]::Compare($a.Rel, $b.Rel, [System.StringComparison]::OrdinalIgnoreCase) }
    $list.Sort($cmp)
    $script:Files = $list
}

# ---------- ③ / ⑥ 一覧つきメニュー（戻り値: 選んだ項目の番号、Esc は -1） ----------
function Show-ListMenu($title, $infoLines, $listHeader, $listItems, $menuItems, $escText = '戻る') {
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
            $range = ('   [{0}-{1} / {2}]  PageUp / PageDown でスクロール' -f ($scroll + 1), ([Math]::Min($n, $scroll + $listRows)), $n)
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
        $lines.Add((mkLine ('  上下キー: 移動    Enter: 決定    Esc: ' + $escText) 'DarkGray'))
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

# ---------- ④ 順番を指定して結合（1 つずつ選ぶ） ----------
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
        $lines.Add((mkLine '順番を指定して結合' 'Yellow'))
        $lines.Add((mkLine '  結合したい順に Enter で 1 つずつ選びます（もう一度 Enter で取り消し）。選ばなかったファイルは結合しません。'))
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
                $text = '=== 決定して次へ（' + $order.Count + ' 件選択中）'
                $fg = 'Cyan'
            }
            elseif ($i -eq $n + 1) {
                $text = '=== やり直し（選択をすべて解除）'
                $fg = 'Cyan'
            }
            else {
                $text = '=== 戻る'
                $fg = 'Cyan'
            }
            if ($i -eq $cur) { $lines.Add((mkLine ('> ' + $text) 'Black' 'Gray')) }
            else { $lines.Add((mkLine ('  ' + $text) $fg)) }
        }
        $lines.Add((mkLine ('-' * $cols) 'DarkGray'))
        $lines.Add((mkLine ('  ' + $msg) 'Yellow'))
        $lines.Add((mkLine '  上下キー: 移動    Enter: 選ぶ / 取り消す    Home / End / PageUp / PageDown: ジャンプ    Esc: 戻る' 'DarkGray'))
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
                if ($order.Count -eq 0) { $msg = '1 件以上選んでから「決定」してください。' }
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

# ---------- ⑤ 設定 ----------
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
        $lines.Add((mkLine '設定' 'Yellow'))
        $lines.Add((mkLine ('  設定ファイル: ' + $script:SettingsPath) 'DarkGray'))
        $lines.Add((mkLine ('-' * $cols) 'DarkGray'))
        for ($i = 0; $i -lt $cnt; $i++) {
            $mark = '[ ]'
            if ($script:Cfg[$defs[$i].Key]) { $mark = '[x]' }
            $text = $mark + ' ' + $defs[$i].Label
            if ($i -eq $cur) { $lines.Add((mkLine ('> ' + $text) 'Black' 'Gray')) }
            else { $lines.Add((mkLine ('  ' + $text))) }
        }
        if ($cur -eq $cnt) { $lines.Add((mkLine '> 戻る' 'Black' 'Gray')) }
        else { $lines.Add((mkLine '  戻る')) }
        $lines.Add((mkLine ('-' * $cols) 'DarkGray'))
        if ($cur -lt $cnt) { $lines.Add((mkLine ('  ' + $defs[$cur].Desc) 'Cyan')) }
        else { $lines.Add((mkLine '')) }
        $lines.Add((mkLine ('  ' + $msg) 'Yellow'))
        $lines.Add((mkLine '  上下キー: 移動    Enter: オン / オフ切り替え（自動保存）    Esc: 戻る' 'DarkGray'))
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
            else { $msg = '設定ファイルを保存できませんでした（今回の起動中のみ有効です）。' }
        }
    }
}

# ---------- Office -> PDF ----------
function Convert-Office($f, [string]$outPdf) {
    $ext = $f.Ext
    if ($ext -eq '.ppt' -or $ext -eq '.pptx') {
        if (-not $script:Apps.ppt) {
            try { $script:Apps.ppt = New-Object -ComObject PowerPoint.Application }
            catch { throw 'PowerPoint を起動できません（インストールされていますか？）' }
        }
        $pres = $script:Apps.ppt.Presentations.Open($f.Full, -1, 0, 0)    # 読み取り専用・ウィンドウなし
        try { $pres.SaveAs($outPdf, 32) }                                  # 32 = PDF
        finally { $pres.Close() }
    }
    elseif ($ext -eq '.doc' -or $ext -eq '.docx') {
        if (-not $script:Apps.word) {
            try { $w = New-Object -ComObject Word.Application }
            catch { throw 'Word を起動できません（インストールされていますか？）' }
            $w.Visible = $false
            try { $w.DisplayAlerts = 0 } catch { }
            try { $w.AutomationSecurity = 3 } catch { }
            $script:Apps.word = $w
        }
        $doc = $script:Apps.word.Documents.Open($f.Full, $false, $true, $false)   # 変換確認なし・読み取り専用・履歴に残さない
        try { $doc.ExportAsFixedFormat($outPdf, 17) }                              # 17 = PDF
        finally { $doc.Close(0) }                                                  # 0 = 保存しない
    }
    elseif ($ext -eq '.xls' -or $ext -eq '.xlsx' -or $ext -eq '.xlsm') {
        if (-not $script:Apps.excel) {
            try { $x = New-Object -ComObject Excel.Application }
            catch { throw 'Excel を起動できません（インストールされていますか？）' }
            $x.Visible = $false
            $x.DisplayAlerts = $false
            try { $x.EnableEvents = $false } catch { }
            try { $x.AutomationSecurity = 3 } catch { }
            try { $x.AskToUpdateLinks = $false } catch { }
            $script:Apps.excel = $x
        }
        $wb = $script:Apps.excel.Workbooks.Open($f.Full, 0, $true)    # リンク更新なし・読み取り専用
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
                    $script:Warns.Add($f.Rel + ' : Excel の印刷設定を変更できないシートがありました（プリンタ未設定の可能性）。元の設定のまま変換しています。')
                }
            }
            $wb.ExportAsFixedFormat(0, $outPdf)                         # 0 = PDF
        }
        finally { $wb.Close($false) }
    }
    else {
        throw ('未対応の形式です: ' + $ext)
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

# ---------- ⑦ 変換・結合 / ⑧ 完了 ----------
function Invoke-Merge($sel, [string]$outPath) {
    $script:RunOk = $false
    $script:Warns.Clear()
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('pdfmerge_' + [Guid]::NewGuid().ToString('N'))
    [Console]::Clear()
    [Console]::CursorVisible = $true
    try {
        Write-Host '変換と結合を始めます...' -ForegroundColor Cyan
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
                    Write-Host ('    スキップ: ' + $err) -ForegroundColor Yellow
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

        if ($okFiles.Count -eq 0) { throw '結合できるファイルがありません。' }
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
        Write-Host '結合しています...'
        $totalPages = [PdfMerge]::Merge($paths, $titles, $blank, [bool]$script:Cfg['UnifyA4'], $skip, $outPath, $start, $errs)

        for ($i = 0; $i -lt $n; $i++) {
            if ($start[$i] -eq 0) {
                $skipped.Add([pscustomobject]@{ Rel = $okFiles[$i].Rel; Reason = $errs[$i] })
            }
        }
        if ($skipped.Count -gt 0) {
            Write-Host ''
            Write-Host ('スキップしたファイル（' + $skipped.Count + ' 件）:') -ForegroundColor Yellow
            foreach ($s in $skipped) { Write-Host ('  ' + $s.Rel + ' : ' + $s.Reason) -ForegroundColor Yellow }
        }
        if ($script:Warns.Count -gt 0) {
            Write-Host ''
            Write-Host '注意:' -ForegroundColor Yellow
            foreach ($w in $script:Warns) { Write-Host ('  ' + $w) -ForegroundColor Yellow }
        }
        Write-Host ''
        Write-Host '結合内容:'
        for ($i = 0; $i -lt $n; $i++) {
            if ($start[$i] -gt 0) {
                $mark = ''
                if ($blank[$i]) { $mark = ' (+白紙1p)' }
                Write-Host (('  p.{0,3} から {1,3}p{2}  {3}' -f $start[$i], $okCounts[$i], $mark, $okFiles[$i].Rel))
            }
        }
        Write-Host ''
        Write-Host ('完了: ' + $outPath + '  （全 ' + $totalPages + ' ページ）') -ForegroundColor Green
        $script:RunOk = $true

        if ($script:Cfg['OpenFolderAfter']) {
            try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"' + $outPath + '"') }
            catch { }
        }
    }
    catch {
        Write-Host ''
        Write-Host ('エラー: ' + (Get-Msg $_)) -ForegroundColor Red
    }
    finally {
        Close-Apps
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath ($outPath + '.tmp')) { Remove-Item -LiteralPath ($outPath + '.tmp') -Force -ErrorAction SilentlyContinue }
    }
}

function On-Off([bool]$v) { if ($v) { return 'オン' } else { return 'オフ' } }

# ---------- メイン ----------
function Main {
    $script:DefFg = [Console]::ForegroundColor
    $script:DefBg = [Console]::BackgroundColor

    try { Add-Type -TypeDefinition $cs -Language CSharp }
    catch { throw ('C# のコンパイルに失敗しました (Add-Type): ' + (Get-Msg $_)) }

    Load-Settings

    # ① フォルダ選択（フォルダを cmd にドラッグした場合はダイアログを省略）
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
    $outName = $leaf + '結合版' + (Get-Date -Format 'yyyyMMdd') + '.pdf'
    $outRegex = '^' + [regex]::Escape($leaf) + '結合版\d{8}\.pdf$'
    $outPath = Join-Path $dir $outName

    while ($true) {
        # ② ファイル収集
        Get-Targets $dir $outRegex
        $files = $script:Files
        if ($files.Count -eq 0) {
            [Console]::Clear()
            Write-Host ('対象のファイル（PDF / PowerPoint / Word / Excel）が見つかりません: ' + $dir) -ForegroundColor Yellow
            Wait-Enter 'Enter キーで終了します。'
            return
        }
        $labels = New-Object 'System.Collections.Generic.List[string]'
        for ($i = 0; $i -lt $files.Count; $i++) { $labels.Add(('{0,3}. {1}' -f ($i + 1), $files[$i].Rel)) }
        $dirShown = [TuiUtil]::CutHead($dir, [Math]::Max(10, (Get-Cols) - 12))
        $info = @(
            ('フォルダ: ' + $dirShown),
            ('対象: ' + $files.Count + ' 件（PDF / PowerPoint / Word / Excel・サブフォルダを含む・ファイル名の昇順）')
        )

        # ③ メインメニュー
        $c = Show-ListMenu 'File2PDF' $info '対象ファイル一覧' $labels @('このまま結合', '順番を指定して結合', '設定', '終了') '終了'
        $sel = $null
        if ($c -eq 0) { $sel = $files }
        elseif ($c -eq 1) {
            Show-OrderSelect $files        # ④
            if ($script:PickOk) { $sel = $script:Picked }
        }
        elseif ($c -eq 2) { Show-Settings; continue }   # ⑤
        elseif ($c -eq 3 -or $c -eq -1) { return }
        else { continue }
        if ($null -eq $sel) { continue }

        # ⑥ 確認
        $exists = Test-Path -LiteralPath $outPath
        $labels2 = New-Object 'System.Collections.Generic.List[string]'
        for ($i = 0; $i -lt $sel.Count; $i++) { $labels2.Add(('{0,3}. {1}' -f ($i + 1), $sel[$i].Rel)) }
        $outShown = [TuiUtil]::CutHead($outPath, [Math]::Max(10, (Get-Cols) - 12))
        $info2 = New-Object 'System.Collections.Generic.List[string]'
        $info2.Add('出力先: ' + $outShown)
        $info2.Add('設定: A4統一=' + (On-Off $script:Cfg['UnifyA4']) + ' / Excel横1ページ=' + (On-Off $script:Cfg['ExcelFitWidth']) + ' / エラー時スキップ=' + (On-Off $script:Cfg['SkipOnError']))
        $runLabel = '実行'
        if ($exists) {
            $info2.Add('※ 同名のファイルが既にあります。実行すると上書きされます。')
            $runLabel = '上書きして実行'
        }
        $c2 = Show-ListMenu '結合内容の確認' $info2 ('結合する順番（' + $sel.Count + ' 件）') $labels2 @($runLabel, '戻る')
        if ($c2 -ne 0) { continue }

        # ⑦ 実行 → ⑧ 完了
        Invoke-Merge $sel $outPath
        if ($script:Cfg['ReturnToMenu']) { Wait-Enter 'Enter キーでメニューに戻ります。' }
        else {
            Wait-Enter 'Enter キーで終了します。'
            return
        }
    }
}

try { Main }
catch {
    [Console]::CursorVisible = $true
    Write-Host ''
    Write-Host ('エラー: ' + (Get-Msg $_)) -ForegroundColor Red
    $script:ExitCode = 1
}
finally {
    try { Close-Apps } catch { }
    try { [Console]::CursorVisible = $true; [Console]::ResetColor() } catch { }
}
exit $script:ExitCode
