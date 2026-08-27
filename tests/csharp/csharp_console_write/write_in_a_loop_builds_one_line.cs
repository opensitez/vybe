// vybe-test: csharp/csharp_console_write/write_in_a_loop_builds_one_line
// origin: languages/csharp/tests/csharp/test_csharp_console_write.rs

using static __Harness;

for (int i = 0; i < 3; i++) { __Pr((i).ToString()); }
__P("");
__Check("012");

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
