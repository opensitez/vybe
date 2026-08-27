// vybe-test: csharp/csharp_property_accessor_matrix/property_accessor_matrix_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_property_accessor_matrix.rs

using static __Harness;

// property_accessor_matrix
int seed = 64;
__P(((seed * 2) / 2 == seed || seed == 0).ToString());
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
