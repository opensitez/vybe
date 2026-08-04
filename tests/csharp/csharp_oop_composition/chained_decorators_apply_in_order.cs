// vybe-test: csharp/csharp_oop_composition/chained_decorators_apply_in_order
// origin: languages/csharp/tests/csharp/test_csharp_oop_composition.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

interface IText{string Get();}
class Plain:IText{public string Get()=>"hello";}
class Shout:IText{IText i;public Shout(IText x){i=x;}public string Get()=>i.Get().ToUpper();}
class Wrap:IText{IText i;public Wrap(IText x){i=x;}public string Get()=>$"[{i.Get()}]";}
IText t=new Wrap(new Shout(new Plain()));
__P((t.Get()).ToString());
__Check("[HELLO]");
