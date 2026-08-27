// vybe-test: csharp/csharp_string_advanced_ops/string_join_with_ienumerable_source
// origin: languages/csharp/tests/csharp/test_csharp_string_advanced_ops.rs

using static __Harness;

var nums=Enumerable.Range(1,5);
__P((string.Join("-",nums)).ToString());
__Check("1-2-3-4-5");

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
