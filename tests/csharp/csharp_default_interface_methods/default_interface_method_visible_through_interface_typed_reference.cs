// vybe-test: csharp/csharp_default_interface_methods/default_interface_method_visible_through_interface_typed_reference
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ICounter {
    int Value { get; }
    int Next() { return Value + 1; }
}
class Counter : ICounter {
    public int Value { get; set; }
}
ICounter counter = new Counter { Value = 4 };
__Check((counter.Next()).ToString(), "5");
