// vybe-test: csharp/csharp_oop_composition/strategy_injected_via_constructor_delegation
// origin: languages/csharp/tests/csharp/test_csharp_oop_composition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IFormatter{string Format(int n);}
class Hex:IFormatter{public string Format(int n)=>n.ToString("X");}
class Dec:IFormatter{public string Format(int n)=>n.ToString();}
class Printer{
    IFormatter _f;
    public Printer(IFormatter f){_f=f;}
    public string Print(int n)=>_f.Format(n);
}
__Check((new Printer(new Hex()).Print(255)).ToString(), "FF");
__Check((new Printer(new Dec()).Print(255)).ToString(), "255");
