// vybe-test: csharp/csharp_loops/goto_jumps_forward_to_labeled_statement
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

using static __Harness;

int x = 0;
goto done;
x = 99;
done:
__P((x).ToString());
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
