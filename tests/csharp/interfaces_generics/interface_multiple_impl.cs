// vybe-test: csharp/interfaces_generics/interface_multiple_impl
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IShape {
    double Area();
}
class Circle : IShape {
    public double Radius;
    public double Area() { return 3.14 * Radius * Radius; }
}
class Square : IShape {
    public double Side;
    public double Area() { return Side * Side; }
}
IShape c = new Circle { Radius = 10 };
IShape s = new Square { Side = 5 };
__Check((c.Area()).ToString(), "314");
__Check((s.Area()).ToString(), "25");
