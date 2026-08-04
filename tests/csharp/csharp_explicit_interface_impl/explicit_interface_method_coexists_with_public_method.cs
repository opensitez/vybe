// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_method_coexists_with_public_method
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

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

interface IFormatter { string Format(); }
class Report : IFormatter {
    public string Format() { return "public"; }
    string IFormatter.Format() { return "explicit"; }
}
var report = new Report();
__P((report.Format()).ToString());
__P((((IFormatter)report).Format()).ToString());
__Check("public\nexplicit");
