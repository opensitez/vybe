// vybe-test: csharp/csharp_console_write/write_integer_then_write_integer_concatenate
// origin: languages/csharp/tests/csharp/test_csharp_console_write.rs

using static __Harness;

// console_write
__Pr((1).ToString());
__Pr((2).ToString());
__Pr((3).ToString());
__P("");
__Check("123");

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
