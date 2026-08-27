// vybe-test: csharp/csharp_local_function_static/local_function_overload_by_param_count
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int Compute(int n){int One(int x)=>x+1; int Two(int x,int y)=>x+y; return Two(n,One(n));}
__P((Compute(5)).ToString());
__Check("11");

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
