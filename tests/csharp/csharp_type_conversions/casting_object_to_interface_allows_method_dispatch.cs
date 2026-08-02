// vybe-test: csharp/csharp_type_conversions/casting_object_to_interface_allows_method_dispatch
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IGreeter { string Say(); } class Greeter : IGreeter { public string Say() { return "hi"; } } object item = new Greeter(); __Check((((IGreeter)item).Say()).ToString(), "hi");
