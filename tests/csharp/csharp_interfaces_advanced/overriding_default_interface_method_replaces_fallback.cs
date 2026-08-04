// vybe-test: csharp/csharp_interfaces_advanced/overriding_default_interface_method_replaces_fallback
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
class Bob:IGreeter{
    public string Name()=>"Bob";
    public string Greet()=>"Hi "+Name()+"!";
}
IGreeter g=new Bob();
__P((g.Greet()).ToString());
__Check("Hi Bob!");
