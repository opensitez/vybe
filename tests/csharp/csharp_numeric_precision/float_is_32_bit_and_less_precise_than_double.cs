// vybe-test: csharp/csharp_numeric_precision/float_is_32_bit_and_less_precise_than_double
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

using static __Harness;

float f=1.0f/3.0f;
double d=1.0/3.0;
__P((f==(float)d).ToString());
__Check("True");

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
