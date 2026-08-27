// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_string_key_get_set
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

var m = new Map();
m["count"] = 7;
__P((m["count"]).ToString());
__Check("7");

class Map { System.Collections.Generic.Dictionary<string, int> d = new(); public int this[string k] { get => d[k]; set => d[k] = value; } }

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
