// vybe-test: csharp/csharp_method_group_delegates/method_group_converts_to_func_without_explicit_lambda_wrapper
// origin: languages/csharp/tests/csharp/test_csharp_method_group_delegates.rs

using static __Harness;

static int Double(int n) => n * 2;
System.Func<int, int> fn = Double;
__P((fn(6)).ToString());
__Check("12");

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
