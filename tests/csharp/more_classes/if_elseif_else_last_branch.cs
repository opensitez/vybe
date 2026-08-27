// vybe-test: csharp/more_classes/if_elseif_else_last_branch
// origin: languages/csharp/tests/csharp/test_more_classes.rs

using static __Harness;

var x = 5;
if (x > 20) { __P(("big").ToString()); }
else if (x > 10) { __P(("medium").ToString()); }
else { __P(("small").ToString()); }
__Check("small");

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
