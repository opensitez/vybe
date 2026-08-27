// vybe-test: csharp/csharp_ranges_indices/range_on_string_returns_substring
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

using static __Harness;

string s="hello world";
__P((s[6..]).ToString());
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
