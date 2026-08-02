// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_method_on_generic_interface_returns_value
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IBox<T> { T Unwrap(); }
class NumberBox : IBox<int> {
    int IBox<int>.Unwrap() { return 42; }
}
IBox<int> box = new NumberBox();
__Check((box.Unwrap()).ToString(), "42");
