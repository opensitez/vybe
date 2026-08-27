// vybe-test: csharp/csharp_string_methods/substring_extracts_region_by_start_and_length
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

using static __Harness;

__P(("hello world".Substring(6, 5)).ToString());
__Check("world");

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
