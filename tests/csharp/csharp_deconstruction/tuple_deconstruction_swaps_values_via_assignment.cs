// vybe-test: csharp/csharp_deconstruction/tuple_deconstruction_swaps_values_via_assignment
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

using static __Harness;

int left = 1;
int right = 2;
(left, right) = (right, left);
__P((left).ToString());
__P((right).ToString());
__Check("2\n1");

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
