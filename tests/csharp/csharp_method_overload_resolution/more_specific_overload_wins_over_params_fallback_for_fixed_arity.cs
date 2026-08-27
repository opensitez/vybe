// vybe-test: csharp/csharp_method_overload_resolution/more_specific_overload_wins_over_params_fallback_for_fixed_arity
// origin: languages/csharp/tests/csharp/test_csharp_method_overload_resolution.rs

using static __Harness;

int res1 = Calculator.Add(5, 5);
double res2 = Calculator.Add(2.5, 2.5);
__P(res1.ToString());
__P(res2.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("10\n5");

class Calculator {
    public static int Add(int a, int b) => a + b;
    public static double Add(double a, double b) => a + b;
}
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
