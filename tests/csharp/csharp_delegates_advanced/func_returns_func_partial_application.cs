// vybe-test: csharp/csharp_delegates_advanced/func_returns_func_partial_application
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

using static __Harness;

System.Func<int,System.Func<int,int>> multiply=factor=>n=>n*factor;
var triple=multiply(3);
__P((triple(7)).ToString());
__Check("21");

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
