// vybe-test: csharp/csharp_control_flow/if_elseif_chain
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

using static __Harness;

int score = 75;
if (score >= 90) __P(("A").ToString());
else if (score >= 80) __P(("B").ToString());
else if (score >= 70) __P(("C").ToString());
else __P(("F").ToString());
__Check("C");

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
