// vybe-test: csharp/csharp_oop_composition/chained_decorators_apply_in_order
// origin: languages/csharp/tests/csharp/test_csharp_oop_composition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IText{string Get();}
class Plain:IText{public string Get()=>"hello";}
class Shout:IText{IText i;public Shout(IText x){i=x;}public string Get()=>i.Get().ToUpper();}
class Wrap:IText{IText i;public Wrap(IText x){i=x;}public string Get()=>$"[{i.Get()}]";}
IText t=new Wrap(new Shout(new Plain()));
__Check((t.Get()).ToString(), "[HELLO]");
