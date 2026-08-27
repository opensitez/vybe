// vybe-test: csharp/csharp_ranges_indices/index_variable_used_in_array_access
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

using static __Harness;

int[] a={10,20,30,40,50}
;
System.Index i=^2;
__P((a[i]).ToString());
__Check("40");

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
