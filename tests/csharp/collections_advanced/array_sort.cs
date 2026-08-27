// vybe-test: csharp/collections_advanced/array_sort
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

int[] arr = { 5, 3, 8, 1, 2 }
;
Array.Sort(arr);
foreach (var x in arr) __P((x).ToString());
__Check("1\n2\n3\n5\n8");

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
