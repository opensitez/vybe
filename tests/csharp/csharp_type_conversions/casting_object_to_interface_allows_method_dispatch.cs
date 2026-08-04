// vybe-test: csharp/csharp_type_conversions/casting_object_to_interface_allows_method_dispatch
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

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

interface IGreeter { string Say(); } class Greeter : IGreeter { public string Say() { return "hi"; } } object item = new Greeter(); __P((((IGreeter)item).Say()).ToString());
__Check("hi");
