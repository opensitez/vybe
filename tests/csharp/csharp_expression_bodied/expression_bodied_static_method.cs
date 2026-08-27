// vybe-test: csharp/csharp_expression_bodied/expression_bodied_static_method
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

using static __Harness;

__P((Utils.Clamp(15,0,10)).ToString());
__Check("10");

static class Utils{public static int Clamp(int v,int lo,int hi)=>v<lo?lo:v>hi?hi:v;}

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
