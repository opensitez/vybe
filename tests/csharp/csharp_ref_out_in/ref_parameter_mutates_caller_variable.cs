// vybe-test: csharp/csharp_ref_out_in/ref_parameter_mutates_caller_variable
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

using static __Harness;

void Double(ref int x){x*=2;}
int n=5;
Double(ref n);
__P((n).ToString());
__Check("10");

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
