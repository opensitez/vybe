// vybe-test: csharp/csharp_goto_switch_labels/goto_label_jumps_to_shared_cleanup
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int n = 1;
string msg = "";
if (n == 1) goto cleanup;
msg = "skip";
cleanup:
msg = "ok";
__P((msg).ToString());
__Check("ok");

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
