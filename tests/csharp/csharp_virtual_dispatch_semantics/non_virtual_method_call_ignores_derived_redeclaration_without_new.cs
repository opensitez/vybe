// vybe-test: csharp/csharp_virtual_dispatch_semantics/non_virtual_method_call_ignores_derived_redeclaration_without_new
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Printer {
    public string Format(int value) { return "p:" + value; }
}
class FancyPrinter : Printer {
    public string Format(int value) { return "f:" + value; }
}
Printer tool = new FancyPrinter();
__Check((tool.Format(7)).ToString(), "p:7");
