// vybe-test: csharp/csharp_loops/while_loop_body_skipped_when_condition_initially_false
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

using static __Harness;

int count=0;
while(false) count++;
__P((count).ToString());
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
