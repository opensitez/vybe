// vybe-test: csharp/csharp_generics_constraints/generic_method_with_multiple_constraints_uses_interface_and_constructor
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IValue { int Read(); } class Item : IValue { public int Read() { return 4; } } int Build<T>() where T : IValue, new() { return new T().Read(); } __Check((Build<Item>()).ToString(), "4");
