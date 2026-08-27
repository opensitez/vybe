// vybe-test: csharp/csharp_interfaces_advanced/default_interface_method_provides_fallback_implementation
// origin: languages/csharp/tests/csharp/test_csharp_interfaces_advanced.rs

using static __Harness;

IGreeter g=new Alice();
__P((g.Greet()).ToString());
__Check("Hello Alice");

interface IGreeter{
    string Name();
    string Greet()=>"Hello "+Name();
}

class Alice:IGreeter{public string Name()=>"Alice";}

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
