// vybe-test: csharp/csharp_console_write/lone_write_without_trailing_newline_is_still_captured
// origin: languages/csharp/tests/csharp/test_csharp_console_write.rs

using static __Harness;

// console_write
__Pr(("solo").ToString());
__Check("solo");

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
