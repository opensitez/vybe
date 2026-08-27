// vybe-test: csharp/csharp_expression_bodied/expression_bodied_operator
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

using static __Harness;

__P(((new Num{V=3}+new Num{V=4}).V).ToString());
__Check("7");

struct Num{public int V;public static Num operator+(Num a,Num b)=>new Num{V=a.V+b.V};}

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
