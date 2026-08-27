// vybe-test: csharp/csharp_array_operations/array_clear_fills_range_with_default_values
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

using static __Harness;

int[] a = {1,2,3,4,5}
;
System.Array.Clear(a, 1, 3);
__P((a[0]).ToString());
__P((a[2]).ToString());
__Check("1\n0");

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
