// vybe-test: csharp/csharp_interface_explicit_impl/explicit_impl_is_not_accessible_through_class_reference
// origin: languages/csharp/tests/csharp/test_csharp_interface_explicit_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IArea { double Area(); }
class Square : IArea {
    public double Side;
    double IArea.Area() => Side * Side;
}
IArea shape = new Square { Side = 3 };
__Check((shape.Area()).ToString(), "9");
