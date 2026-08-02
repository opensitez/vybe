// vybe-test: csharp/csharp_interfaces_advanced/default_interface_method_provides_fallback_implementation
// origin: languages/csharp/tests/csharp/test_csharp_interfaces_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IGreeter{
    string Name();
    string Greet()=>"Hello "+Name();
}
class Alice:IGreeter{public string Name()=>"Alice";}
IGreeter g=new Alice();
__Check((g.Greet()).ToString(), "Hello Alice");
