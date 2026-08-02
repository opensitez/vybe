// vybe-test: csharp/csharp_oop_composition/decorator_wraps_and_extends_base_behaviour
// origin: languages/csharp/tests/csharp/test_csharp_oop_composition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IText{string Get();}
class Plain:IText{public string Get()=>"hello";}
class Shout:IText{
    IText _inner;
    public Shout(IText inner){_inner=inner;}
    public string Get()=>_inner.Get().ToUpper()+"!";
}
IText t=new Shout(new Plain());
__Check((t.Get()).ToString(), "HELLO!");
