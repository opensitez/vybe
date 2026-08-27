// vybe-test: csharp/csharp_console_write/write_emits_no_newline_so_consecutive_writes_share_a_line
// origin: languages/csharp/tests/csharp/test_csharp_console_write.rs

using static __Harness;

// console_write
__Pr(("a").ToString());
__Pr(("b").ToString());
__P("");
__Check("ab");

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
