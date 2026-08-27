// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

using static __Harness;

// for_loop_bounds
var map = new System.Collections.Generic.Dictionary<int, int>();
map[45] = 46;
__P((map.ContainsKey(45) && map[45] == 46).ToString());
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
