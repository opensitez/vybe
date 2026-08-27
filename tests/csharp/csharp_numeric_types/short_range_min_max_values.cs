// vybe-test: csharp/csharp_numeric_types/short_range_min_max_values
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

using static __Harness;

__P((short.MinValue).ToString());
__P((short.MaxValue).ToString());
__Check("-32768\n32767");

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
