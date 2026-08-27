// vybe-test: csharp/csharp_local_functions/local_function_captures_outer_variable
// origin: languages/csharp/tests/csharp/test_csharp_local_functions.rs

using static __Harness;

int multiplier=3;
int Mul(int n){
    int Scaled(int x)=>x*multiplier;
    return Scaled(n);
}
__P((Mul(7)).ToString());
__Check("21");

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
