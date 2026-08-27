// vybe-test: csharp/csharp_local_function_static/local_function_captures_outer_int
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int offset=10;
int Add(int n){int B(int x)=>x+offset; return B(n);}
__P((Add(5)).ToString());
__Check("15");

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
