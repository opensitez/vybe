// vybe-test: csharp/csharp_namespace_aliases/using_alias_can_target_generic_type
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;
using TextList = System.Collections.Generic.List<string>;

var list = new TextList { "a", "b" }
;
__P((list.Count).ToString());
__Check("2");

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
