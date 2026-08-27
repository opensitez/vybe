// vybe-test: csharp/csharp_local_function_static/local_function_recursive_factorial
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int Fact(int n){int F(int k)=>k<=1?1:k*F(k-1); return F(n);}
__P((Fact(5)).ToString());
__Check("120");

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
