// vybe-test: csharp/csharp_local_function_static/local_function_with_default_parameter
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int Inc(int n){int Step(int x,int by=1)=>x+by; return Step(n,3);}
__P((Inc(10)).ToString());
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
