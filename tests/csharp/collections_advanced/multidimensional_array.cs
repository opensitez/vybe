// vybe-test: csharp/collections_advanced/multidimensional_array
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

int[,] matrix = { { 1, 2 }, { 3, 4 }, { 5, 6 } }
;
__P((matrix[0, 0]).ToString());
__P((matrix[1, 1]).ToString());
__P((matrix[2, 0]).ToString());
__Check("1\n4\n5");

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
