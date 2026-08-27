// vybe-test: csharp/csharp_linq_projections/skip_while_skips_until_predicate_fails
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

using static __Harness;

var result = new[]{1,2,3,4,5}
.SkipWhile(x => x<3);
foreach(var n in result) __P((n).ToString());
__Check("3\n4\n5");

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
