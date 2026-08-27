// vybe-test: csharp/csharp_numeric_types/hex_integer_literal_parsed_correctly
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

using static __Harness;

int n = 0xFF;
__P((n).ToString());
__Check("255");

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
