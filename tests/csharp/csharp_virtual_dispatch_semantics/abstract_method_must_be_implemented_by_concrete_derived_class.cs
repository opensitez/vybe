// vybe-test: csharp/csharp_virtual_dispatch_semantics/abstract_method_must_be_implemented_by_concrete_derived_class
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

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

abstract class Parser {
    public abstract string Parse(string input);
}
class EchoParser : Parser {
    public override string Parse(string input) { return input.Trim(); }
}
Parser parser = new EchoParser();
__P((parser.Parse("  hi  ")).ToString());
__Check("hi");
