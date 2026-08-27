// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

using static __Harness;

// jagged_array_patterns
var map = new System.Collections.Generic.Dictionary<int, int>();
map[28] = 29;
__P((map.ContainsKey(28) && map[28] == 29).ToString());
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
