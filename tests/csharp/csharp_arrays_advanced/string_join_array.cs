// vybe-test: csharp/csharp_arrays_advanced/string_join_array
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

using static __Harness;

var arr = new[] { "a", "b", "c" }
;
__P((string.Join(",", arr)).ToString());
__P((string.Join(" - ", arr)).ToString());
__Check("a,b,c\na - b - c");

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
