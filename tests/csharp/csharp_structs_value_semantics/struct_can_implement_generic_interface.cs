// vybe-test: csharp/csharp_structs_value_semantics/struct_can_implement_generic_interface
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IBox<T> { T Read(); } struct NumberBox : IBox<int> { public int Read() { return 14; } } IBox<int> box = new NumberBox(); __Check((box.Read()).ToString(), "14");
