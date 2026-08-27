// vybe-test: csharp/csharp_random_distribution_matrix/random_distribution_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_random_distribution_matrix.rs

using static __Harness;

// random_distribution_matrix
var map = new System.Collections.Generic.Dictionary<int, int>();
map[98] = 99;
__P((map.ContainsKey(98) && map[98] == 99).ToString());
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
