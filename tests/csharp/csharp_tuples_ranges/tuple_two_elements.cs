// vybe-test: csharp/csharp_tuples_ranges/tuple_two_elements
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

using static __Harness;

var t = (10, "hello");
__P((t.Item1).ToString());
__P((t.Item2).ToString());
__Check("10\nhello");

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
