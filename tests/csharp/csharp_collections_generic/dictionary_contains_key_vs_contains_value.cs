// vybe-test: csharp/csharp_collections_generic/dictionary_contains_key_vs_contains_value
// origin: languages/csharp/tests/csharp/test_csharp_collections_generic.rs

using static __Harness;

var d=new System.Collections.Generic.Dictionary<string,int>{{"a",1}}
;
__P((d.ContainsKey("a")).ToString());
__P((d.ContainsValue(1)).ToString());
__P((d.ContainsKey("b")).ToString());
__Check("True\nTrue\nFalse");

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
