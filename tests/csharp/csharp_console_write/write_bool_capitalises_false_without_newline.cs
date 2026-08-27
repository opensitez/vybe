// vybe-test: csharp/csharp_console_write/write_bool_capitalises_false_without_newline
// origin: languages/csharp/tests/csharp/test_csharp_console_write.rs

using static __Harness;

// console_write
__Pr((false).ToString());
__P(("!").ToString());
__Check("False!");

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
