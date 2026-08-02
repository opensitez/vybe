// vybe-test: csharp/csharp_primary_constructors/primary_constructor_interface_method_uses_param
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IVal { int Get(); }
class Impl(int n) : IVal { public int Get() => n; }
IVal v = new Impl(12);
__Check((v.Get()).ToString(), "12");
