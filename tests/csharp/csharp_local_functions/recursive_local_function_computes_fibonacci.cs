// vybe-test: csharp/csharp_local_functions/recursive_local_function_computes_fibonacci
// origin: languages/csharp/tests/csharp/test_csharp_local_functions.rs

using static __Harness;

int Fib(int n){
    int F(int k)=>k<=1?k:F(k-1)+F(k-2);
    return F(n);
}
__P((Fib(7)).ToString());
__Check("13");

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
