// vybe-test: csharp/csharp_generics_constraints/generic_interface_implementation_preserves_type_argument
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IBox<T> { T Read(); } class NumberBox : IBox<int> { public int Read() { return 8; } } __Check((((IBox<int>)new NumberBox()).Read()).ToString(), "8");
