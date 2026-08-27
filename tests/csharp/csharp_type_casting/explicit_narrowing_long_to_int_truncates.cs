// vybe-test: csharp/csharp_type_casting/explicit_narrowing_long_to_int_truncates
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

using static __Harness;

long x = 5L;
int y = (int)x;
__P((y).ToString());
__Check("5");

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
