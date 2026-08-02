// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_property_is_visible_through_interface
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IValueHolder { int Value { get; } }
class Counter : IValueHolder {
    int IValueHolder.Value { get { return 12; } }
}
IValueHolder holder = new Counter();
__Check((holder.Value).ToString(), "12");
