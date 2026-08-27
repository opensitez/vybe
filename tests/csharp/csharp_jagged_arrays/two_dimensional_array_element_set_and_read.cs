// vybe-test: csharp/csharp_jagged_arrays/two_dimensional_array_element_set_and_read
// origin: languages/csharp/tests/csharp/test_csharp_jagged_arrays.rs

using static __Harness;

int[,] m = new int[2,2];
m[0,1] = 7;
__P((m[0,1]).ToString());
__Check("7");

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
