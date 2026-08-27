// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

using static __Harness;

// datetime_construction_matrix
var map = new System.Collections.Generic.Dictionary<int, int>();
map[94] = 95;
__P((map.ContainsKey(94) && map[94] == 95).ToString());
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
