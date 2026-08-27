// vybe-test: csharp/csharp_nullable/null_check_with_if
// origin: languages/csharp/tests/csharp/test_csharp_nullable.rs

using static __Harness;

string s = null;
if (s == null) {
    __P(("is null").ToString());
}
else {
    __P(("has value").ToString());
}
s = "test";
if (s != null) {
    __P(("has value").ToString());
}
__Check("is null\nhas value");

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
