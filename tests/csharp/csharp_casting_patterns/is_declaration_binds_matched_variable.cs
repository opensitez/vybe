// vybe-test: csharp/csharp_casting_patterns/is_declaration_binds_matched_variable
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

using static __Harness;

object o=42;
if(o is int n) __P((n*2).ToString());
__Check("84");

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
