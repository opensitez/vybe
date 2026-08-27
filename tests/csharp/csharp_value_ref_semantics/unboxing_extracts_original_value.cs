// vybe-test: csharp/csharp_value_ref_semantics/unboxing_extracts_original_value
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

using static __Harness;

object o=42;
int n=(int)o;
__P((n).ToString());
__Check("42");

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
