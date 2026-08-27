// vybe-test: csharp/csharp_oop_composition/chained_decorators_apply_in_order
// origin: languages/csharp/tests/csharp/test_csharp_oop_composition.rs

using static __Harness;

IText t=new Wrap(new Shout(new Plain()));
__P((t.Get()).ToString());
__Check("[HELLO]");

interface IText{string Get();}

class Plain:IText{public string Get()=>"hello";}

class Shout:IText{IText i;public Shout(IText x){i=x;}public string Get()=>i.Get().ToUpper();}

class Wrap:IText{IText i;public Wrap(IText x){i=x;}public string Get()=>$"[{i.Get()}]";}

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
