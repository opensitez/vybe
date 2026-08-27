// vybe-test: csharp/csharp_delegate_invocation_matrix/delegate_invocation_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_delegate_invocation_matrix.rs

using static __Harness;

// delegate_invocation_matrix
var map = new System.Collections.Generic.Dictionary<int, int>();
map[74] = 75;
__P((map.ContainsKey(74) && map[74] == 75).ToString());
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
