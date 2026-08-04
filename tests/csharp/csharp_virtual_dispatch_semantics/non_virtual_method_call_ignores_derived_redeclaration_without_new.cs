// vybe-test: csharp/csharp_virtual_dispatch_semantics/non_virtual_method_call_ignores_derived_redeclaration_without_new
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

class Printer {
    public string Format(int value) { return "p:" + value; }
}
class FancyPrinter : Printer {
    public string Format(int value) { return "f:" + value; }
}
Printer tool = new FancyPrinter();
__P((tool.Format(7)).ToString());
__Check("p:7");
