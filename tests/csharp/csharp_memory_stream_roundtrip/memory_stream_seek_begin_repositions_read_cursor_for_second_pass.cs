// vybe-test: csharp/csharp_memory_stream_roundtrip/memory_stream_seek_begin_repositions_read_cursor_for_second_pass
// origin: languages/csharp/tests/csharp/test_csharp_memory_stream_roundtrip.rs

using static __Harness;
using System.IO;

var stream = new MemoryStream();
var writer = new StreamWriter(stream);
writer.Write("ab");
writer.Flush();
stream.Seek(0, SeekOrigin.Begin);
__P((stream.ReadByte()).ToString());
stream.Seek(0, SeekOrigin.Begin);
__P((stream.ReadByte()).ToString());
__Check("97\n97");

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
