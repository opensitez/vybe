// vybe-test: csharp/csharp_io_stream/binary_writer_reader_roundtrips_int32
// origin: languages/csharp/tests/csharp/test_csharp_io_stream.rs

using static __Harness;
using var ms = new System.IO.MemoryStream();
using var br = new System.IO.BinaryReader(ms);

using(var bw = new System.IO.BinaryWriter(ms, System.Text.Encoding.UTF8, leaveOpen:true))
    bw.Write(12345);
ms.Position = 0;
__P((br.ReadInt32()).ToString());
__Check("12345");

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
