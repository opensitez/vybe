// vybe-test: csharp/csharp_oop_composition/decorator_wraps_and_extends_base_behaviour
// origin: languages/csharp/tests/csharp/test_csharp_oop_composition.rs

using static __Harness;

IText t=new Shout(new Plain());
__P((t.Get()).ToString());
__Check("HELLO!");

interface IText{string Get();}

class Plain:IText{public string Get()=>"hello";}

class Shout:IText{
    IText _inner;
    public Shout(IText inner){_inner=inner;}
    public string Get()=>_inner.Get().ToUpper()+"!";
}

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
