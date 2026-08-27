// vybe-test: csharp/csharp_array_operations/array_fill_sets_all_elements_to_given_value
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

using static __Harness;

int[] a = new int[4];
System.Array.Fill(a, 7);
__P((a[3]).ToString());
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
