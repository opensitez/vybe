// vybe-test: csharp/collections_advanced/array_exists
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

int[] arr = { 1, 2, 3, 4, 5 }
;
__P((Array.Exists(arr, x => x > 4)).ToString());
__P((Array.Exists(arr, x => x > 10)).ToString());
__Check("True\nFalse");

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
