// vybe-test: csharp/collections_advanced/array_indexof
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

string[] arr = { "a", "b", "c", "d" }
;
__P((Array.IndexOf(arr, "c")).ToString());
__P((Array.IndexOf(arr, "z")).ToString());
__Check("2\n-1");

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
