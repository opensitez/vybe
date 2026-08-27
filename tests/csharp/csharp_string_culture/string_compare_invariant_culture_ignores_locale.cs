// vybe-test: csharp/csharp_string_culture/string_compare_invariant_culture_ignores_locale
// origin: languages/csharp/tests/csharp/test_csharp_string_culture.rs

using static __Harness;

int r=string.Compare("hello","HELLO",System.StringComparison.InvariantCultureIgnoreCase);
__P((r==0).ToString());
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
