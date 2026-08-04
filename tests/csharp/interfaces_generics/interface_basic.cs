// vybe-test: csharp/interfaces_generics/interface_basic
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

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

interface IGreeter {
    string Greet();
}
class HelloGreeter : IGreeter {
    public string Greet() { return "Hello!"; }
}
IGreeter g = new HelloGreeter();
__P((g.Greet()).ToString());
__Check("Hello!");
