// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_non_empty_is_not_empty
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

using static __Harness;

var bytes="x"u8;
__P((bytes.IsEmpty).ToString());
__Check("False");

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
