// vybe-test: csharp/csharp_interface_explicit_impl/explicit_impl_is_not_accessible_through_class_reference
// origin: languages/csharp/tests/csharp/test_csharp_interface_explicit_impl.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

interface IArea { double Area(); }
class Square : IArea {
    public double Side;
    double IArea.Area() => Side * Side;
}
IArea shape = new Square { Side = 3 };
__P((shape.Area()).ToString());
__Check("9");
