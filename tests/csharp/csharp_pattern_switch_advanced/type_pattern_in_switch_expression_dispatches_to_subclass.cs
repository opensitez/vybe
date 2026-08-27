// vybe-test: csharp/csharp_pattern_switch_advanced/type_pattern_in_switch_expression_dispatches_to_subclass
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_advanced.rs

using static __Harness;

int Eval(Expr e)=>e switch{
    Num n=>n.V,
    Add a=>Eval(a.L)+Eval(a.R),
    _=>throw new System.Exception()}
;
var tree=new Add{L=new Num{V=3},R=new Add{L=new Num{V=4},R=new Num{V=5}}}
;
__P((Eval(tree)).ToString());
__Check("12");

abstract class Expr{}

class Num:Expr{public int V;}

class Add:Expr{public Expr L,R;}

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
