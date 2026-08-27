// vybe-test: csharp/csharp_hashset_membership_matrix/hashset_membership_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_hashset_membership_matrix.rs

using static __Harness;

// hashset_membership_matrix
var map = new System.Collections.Generic.Dictionary<int, int>();
map[33] = 34;
__P((map.ContainsKey(33) && map[33] == 34).ToString());
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
