// vybe-test: csharp/csharp_dictionary_operations/try_get_value_returns_false_and_default_on_miss
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

using static __Harness;

var d = new System.Collections.Generic.Dictionary<string,int>();
__P((d.TryGetValue("nope", out int v)).ToString());
__P((v).ToString());
__Check("False\n0");

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
