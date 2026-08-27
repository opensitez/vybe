// vybe-test: csharp/csharp_local_function_static/static_local_function_no_capture_multiply
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int Product(int a,int b){static int Mul(int x,int y)=>x*y; return Mul(a,b);}
__P((Product(6,7)).ToString());
__Check("42");

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
