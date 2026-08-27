// vybe-test: csharp/control_flow/if_elseif_chain_multiple
// origin: languages/csharp/tests/csharp/test_control_flow.rs

using static __Harness;

var x = 2;
if (x == 1) { __P(("one").ToString()); }
else if (x == 2) { __P(("two").ToString()); }
else if (x == 3) { __P(("three").ToString()); }
else { __P(("other").ToString()); }
__Check("two");

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
