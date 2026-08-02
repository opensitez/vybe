// vybe-test: csharp/csharp_interfaces_advanced/overriding_default_interface_method_replaces_fallback
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
class Bob:IGreeter{
    public string Name()=>"Bob";
    public string Greet()=>"Hi "+Name()+"!";
}
IGreeter g=new Bob();
__Check((g.Greet()).ToString(), "Hi Bob!");
