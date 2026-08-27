// vybe-test: csharp/csharp_interlocked_atomic_matrix/interlocked_atomic_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic_matrix.rs

using static __Harness;

// interlocked_atomic_matrix
var map = new System.Collections.Generic.Dictionary<int, int>();
map[83] = 84;
__P((map.ContainsKey(83) && map[83] == 84).ToString());
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
