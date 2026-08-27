// vybe-test: csharp/csharp_using_static/using_static_console_allows_unqualified_writeline
// origin: languages/csharp/tests/csharp/test_csharp_using_static.rs

using static __Harness;
using static System.Console;

WriteLine("hello");

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
