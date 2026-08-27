// vybe-test: csharp/advanced/cast_double_to_int
// origin: languages/csharp/tests/csharp/test_advanced.rs

using static __Harness;

double d = 3.14;
__P((d).ToString());
__Check("3.14");

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
