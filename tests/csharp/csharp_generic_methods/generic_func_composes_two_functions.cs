// vybe-test: csharp/csharp_generic_methods/generic_func_composes_two_functions
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

using static __Harness;

System.Func<A,C> Compose<A,B,C>(System.Func<A,B> f,System.Func<B,C> g)=>x=>g(f(x));
var fn=Compose((int x)=>x*2,(int y)=>y+1);
__P((fn(5)).ToString());
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
