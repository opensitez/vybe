// vybe-test: csharp/csharp_interpolation_basic_matrix/interpolation_basic_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_basic_matrix.rs

using static __Harness;

// interpolation_basic_matrix
int seed = 112;
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
