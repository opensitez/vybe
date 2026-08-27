// vybe-test: csharp/csharp_array_advanced/array_true_for_all_validates_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

using static __Harness;

int[] arr={2,4,6,8}
;
__P((System.Array.TrueForAll(arr,n=>n%2==0)).ToString());
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
