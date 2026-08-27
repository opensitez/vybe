// vybe-test: csharp/csharp_nullable_pattern_matching_matrix/nullable_pattern_matching_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_nullable_pattern_matching_matrix.rs

using static __Harness;

// nullable_pattern_matching_matrix
var map = new System.Collections.Generic.Dictionary<int, int>();
map[125] = 126;
__P((map.ContainsKey(125) && map[125] == 126).ToString());
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
