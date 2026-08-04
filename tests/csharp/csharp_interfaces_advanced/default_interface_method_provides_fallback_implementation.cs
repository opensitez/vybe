// vybe-test: csharp/csharp_interfaces_advanced/default_interface_method_provides_fallback_implementation
// origin: languages/csharp/tests/csharp/test_csharp_interfaces_advanced.rs

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

interface IGreeter{
    string Name();
    string Greet()=>"Hello "+Name();
}
class Alice:IGreeter{public string Name()=>"Alice";}
IGreeter g=new Alice();
__P((g.Greet()).ToString());
__Check("Hello Alice");
