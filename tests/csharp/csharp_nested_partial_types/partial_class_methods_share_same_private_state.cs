// vybe-test: csharp/csharp_nested_partial_types/partial_class_methods_share_same_private_state
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

partial class Counter {
    int value;
}
partial class Counter {
    public void Bump() { value++; }
    public int Read() { return value; }
}
var counter = new Counter();
counter.Bump();
counter.Bump();
__Check((counter.Read()).ToString(), "2");
