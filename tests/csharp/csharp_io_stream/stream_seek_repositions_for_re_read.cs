// vybe-test: csharp/csharp_io_stream/stream_seek_repositions_for_re_read
// origin: languages/csharp/tests/csharp/test_csharp_io_stream.rs

using static __Harness;
using var ms = new System.IO.MemoryStream();

ms.WriteByte(7);
ms.Seek(0, System.IO.SeekOrigin.Begin);
__P((ms.ReadByte()).ToString());
__Check("7");

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
