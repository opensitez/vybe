// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_method_coexists_with_public_method
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IFormatter { string Format(); }
class Report : IFormatter {
    public string Format() { return "public"; }
    string IFormatter.Format() { return "explicit"; }
}
var report = new Report();
__Check((report.Format()).ToString(), "public");
__Check((((IFormatter)report).Format()).ToString(), "explicit");
