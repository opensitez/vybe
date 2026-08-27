// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_range_slice_length
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

using static __Harness;

var bytes="abcdef"u8;
var slice=bytes[2..5];
__P((slice.Length).ToString());
__Check("3");

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
