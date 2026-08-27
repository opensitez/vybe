// vybe-test: csharp/csharp_nullable_pattern_matching_matrix/nullable_pattern_matching_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_nullable_pattern_matching_matrix.rs

using static __Harness;

// nullable_pattern_matching_matrix
int seed = 125;
int[] numbers = new int[] { seed, seed + 1, seed + 2 }
;
__P((numbers.Length == 3).ToString());
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
