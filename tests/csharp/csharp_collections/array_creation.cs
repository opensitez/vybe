// vybe-test: csharp/csharp_collections/array_creation
// origin: languages/csharp/tests/csharp/test_csharp_collections.rs

using static __Harness;

int[] arr = {5, 10, 15, 20, 25}
;
__P((arr.Length).ToString());
__P((arr[2]).ToString());
__Check("5\n15");

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
