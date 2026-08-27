// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_starts_with_prefix_bytes
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

using static __Harness;

var bytes="prefix"u8;
__P((bytes.StartsWith("pre"u8)).ToString());
__Check("True");

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
