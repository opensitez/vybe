// vybe-test: csharp/csharp_error_handling/try_finally
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

using static __Harness;

try {
    __P(("try").ToString());
}
finally {
    __P(("finally").ToString());
}
__Check("try\nfinally");

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
