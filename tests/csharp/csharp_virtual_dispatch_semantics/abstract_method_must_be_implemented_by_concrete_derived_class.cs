// vybe-test: csharp/csharp_virtual_dispatch_semantics/abstract_method_must_be_implemented_by_concrete_derived_class
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((parser.Parse("  hi  ")).ToString(), "hi");
