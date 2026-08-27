// vybe-test: csharp/csharp_string_culture/ordinal_comparison_is_byte_by_byte
// origin: languages/csharp/tests/csharp/test_csharp_string_culture.rs

using static __Harness;

int r=string.CompareOrdinal("a","A");
__P((r>0).ToString());
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
