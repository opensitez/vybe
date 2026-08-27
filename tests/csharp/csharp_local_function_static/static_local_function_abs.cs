// vybe-test: csharp/csharp_local_function_static/static_local_function_abs
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int Abs(int n){static int Pos(int x)=>x<0?-x:x; return Pos(n);}
__P((Abs(-8)).ToString());
__Check("8");

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
