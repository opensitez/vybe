// vybe-test: csharp/csharp_linq_skip_take_distinct/skip_while_on_empty_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

using static __Harness;

var r=System.Array.Empty<int>().SkipWhile(x=>x<1);
__P((r.Count()).ToString());
__Check("0");

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
