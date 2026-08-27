// vybe-test: csharp/csharp_array_operations/array_index_of_returns_first_matching_position
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

using static __Harness;

string[] a = {"a","b","c","b"}
;
__P((System.Array.IndexOf(a,"b")).ToString());
__Check("1");

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
