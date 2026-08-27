// vybe-test: csharp/csharp_interpolation_verbatim_matrix/interpolation_verbatim_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_verbatim_matrix.rs

using static __Harness;

// interpolation_verbatim_matrix
var set = new System.Collections.Generic.HashSet<int>();
set.Add(110);
set.Add(110);
__P((set.Count == 1).ToString());
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
