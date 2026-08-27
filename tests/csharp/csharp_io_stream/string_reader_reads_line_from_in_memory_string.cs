// vybe-test: csharp/csharp_io_stream/string_reader_reads_line_from_in_memory_string
// origin: languages/csharp/tests/csharp/test_csharp_io_stream.rs

using static __Harness;
using var sr = new System.IO.StringReader("line one\nline two");

__P((sr.ReadLine()).ToString());
__Check("line one");

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
