// vybe-test: csharp/csharp_io_path_matrix/io_path_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_io_path_matrix.rs

using static __Harness;

// io_path_matrix
int seed = 122;
__P((seed + 1 > seed).ToString());
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
