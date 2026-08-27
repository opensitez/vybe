// vybe-test: csharp/csharp_while_loop_exit/while_loop_exit_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_while_loop_exit.rs

using static __Harness;

// while_loop_exit
decimal amount = 10m;
__P(((amount / 2m) * 2m == 10m).ToString());
__Check("True");

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
