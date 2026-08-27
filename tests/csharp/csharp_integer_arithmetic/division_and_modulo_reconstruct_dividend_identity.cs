// vybe-test: csharp/csharp_integer_arithmetic/division_and_modulo_reconstruct_dividend_identity
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

using static __Harness;

int dividend = 17;
int divisor = 5;
__P((dividend / divisor * divisor + dividend % divisor).ToString());
__Check("17");

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
