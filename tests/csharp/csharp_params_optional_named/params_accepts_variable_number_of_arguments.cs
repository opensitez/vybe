// vybe-test: csharp/csharp_params_optional_named/params_accepts_variable_number_of_arguments
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

using static __Harness;

int Sum(params int[] ns){int s=0;foreach(var n in ns)s+=n;return s;}
__P((Sum(1,2,3,4,5)).ToString());
__Check("15");

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
