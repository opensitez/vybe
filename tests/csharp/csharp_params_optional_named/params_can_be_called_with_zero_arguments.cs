// vybe-test: csharp/csharp_params_optional_named/params_can_be_called_with_zero_arguments
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

using static __Harness;

int Sum(params int[] ns){int s=0;foreach(var n in ns)s+=n;return s;}
__P((Sum()).ToString());
__Check("0");

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
