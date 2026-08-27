// vybe-test: csharp/csharp_type_casting/int_to_double_implicit_widens_without_loss
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

using static __Harness;

int i = 5;
double d = i;
__P((d).ToString());
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
