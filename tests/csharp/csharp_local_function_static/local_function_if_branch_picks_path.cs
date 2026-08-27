// vybe-test: csharp/csharp_local_function_static/local_function_if_branch_picks_path
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

string Sign(int n){string Pos(int x)=>"+"; string Neg(int x)=>"-"; if(n>=0){return Pos(n);} return Neg(n);}
__P((Sign(-1)).ToString());
__Check("-");

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
