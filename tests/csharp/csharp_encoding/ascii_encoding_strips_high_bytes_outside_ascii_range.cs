// vybe-test: csharp/csharp_encoding/ascii_encoding_strips_high_bytes_outside_ascii_range
// origin: languages/csharp/tests/csharp/test_csharp_encoding.rs

using static __Harness;

var bytes = System.Text.Encoding.ASCII.GetBytes("ABC");
__P((bytes[0]).ToString());
__Check("65");

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
