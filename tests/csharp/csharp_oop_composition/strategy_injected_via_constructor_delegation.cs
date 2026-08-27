// vybe-test: csharp/csharp_oop_composition/strategy_injected_via_constructor_delegation
// origin: languages/csharp/tests/csharp/test_csharp_oop_composition.rs

using static __Harness;

__P((new Printer(new Hex()).Print(255)).ToString());
__P((new Printer(new Dec()).Print(255)).ToString());
__Check("FF\n255");

interface IFormatter{string Format(int n);}

class Hex:IFormatter{public string Format(int n)=>n.ToString("X");}

class Dec:IFormatter{public string Format(int n)=>n.ToString();}

class Printer{
    IFormatter _f;
    public Printer(IFormatter f){_f=f;}
    public string Print(int n)=>_f.Format(n);
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
