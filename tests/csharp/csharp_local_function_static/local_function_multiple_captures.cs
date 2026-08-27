// vybe-test: csharp/csharp_local_function_static/local_function_multiple_captures
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int a=2;
int b=3;
int Mix(int n){int M(int x)=>a*b+x; return M(n);}
__P((Mix(4)).ToString());
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
