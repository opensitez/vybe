// vybe-test: csharp/csharp_primary_constructors/primary_constructor_string_param_used_in_method
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

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

class Greeter(string prefix) {
    public string Greet(string name) => prefix + " " + name;
}
__P((new Greeter("Hello").Greet("World")).ToString());
__Check("Hello World");
