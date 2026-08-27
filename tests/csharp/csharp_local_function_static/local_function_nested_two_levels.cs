// vybe-test: csharp/csharp_local_function_static/local_function_nested_two_levels
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int Outer(int n){int Mid(int x){int Inner(int y)=>y+1; return Inner(x);} return Mid(n);}
__P((Outer(9)).ToString());
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
