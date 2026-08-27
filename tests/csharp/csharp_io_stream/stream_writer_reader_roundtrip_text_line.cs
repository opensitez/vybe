// vybe-test: csharp/csharp_io_stream/stream_writer_reader_roundtrip_text_line
// origin: languages/csharp/tests/csharp/test_csharp_io_stream.rs

using static __Harness;
using var ms = new System.IO.MemoryStream();
using var sr = new System.IO.StreamReader(ms);

using(var sw = new System.IO.StreamWriter(ms, leaveOpen:true)) sw.WriteLine("hello");
ms.Position = 0;
__P((sr.ReadLine()).ToString());
__Check("hello");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
