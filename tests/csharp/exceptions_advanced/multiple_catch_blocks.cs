// vybe-test: csharp/exceptions_advanced/multiple_catch_blocks
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

try {
    throw new ArgumentException("bad arg");
}
catch (ArgumentNullException) {
    __P(("null").ToString());
}
catch (ArgumentException e) {
    __P(("arg: " + e.Message).ToString());
}
catch (Exception) {
    __P(("generic").ToString());
}
__Check("arg: bad arg");

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
