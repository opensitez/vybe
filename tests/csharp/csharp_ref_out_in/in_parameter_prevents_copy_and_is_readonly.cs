// vybe-test: csharp/csharp_ref_out_in/in_parameter_prevents_copy_and_is_readonly
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

using static __Harness;

int Sum3(in int a, in int b, in int c) => a+b+c;
__P((Sum3(1,2,3)).ToString());
__Check("6");

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
