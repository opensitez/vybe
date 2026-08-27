// vybe-test: csharp/csharp_string_split_join/string_split_join_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

using static __Harness;

// string_split_join
var map = new System.Collections.Generic.Dictionary<int, int>();
map[21] = 22;
__P((map.ContainsKey(21) && map[21] == 22).ToString());
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
