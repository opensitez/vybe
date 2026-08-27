// vybe-test: csharp/csharp_memory_stream_roundtrip/memory_stream_write_then_read_returns_same_bytes
// origin: languages/csharp/tests/csharp/test_csharp_memory_stream_roundtrip.rs

using static __Harness;
using System.IO;

var stream = new MemoryStream();
var writer = new StreamWriter(stream);
writer.Write("payload");
writer.Flush();
stream.Position = 0;
var reader = new StreamReader(stream);
__P((reader.ReadToEnd()).ToString());
__Check("payload");

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
