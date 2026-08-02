// vybe-test: csharp/csharp_virtual_dispatch_semantics/interface_method_implementation_is_invoked_through_interface_reference
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IFormatter {
    string Format(int value);
}
class DecimalFormatter : IFormatter {
    public string Format(int value) { return value.ToString("D3"); }
}
IFormatter formatter = new DecimalFormatter();
__Check((formatter.Format(4)).ToString(), "004");
