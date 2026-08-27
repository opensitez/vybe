// vybe-test: csharp/csharp_integer_arithmetic/pre_and_post_increment_on_different_variables
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

using static __Harness;

int left = 2;
int right = 5;
__P((++left + right++).ToString());
__P((left).ToString());
__P((right).ToString());
__Check("8\n3\n6");

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
