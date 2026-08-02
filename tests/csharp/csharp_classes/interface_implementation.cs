// vybe-test: csharp/csharp_classes/interface_implementation
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

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
    public Circle(double r) { Radius = r; }
    public double Area() { return 3.14159 * Radius * Radius; }
}
var c = new Circle(5);
__Check((c.Area()).ToString(), "78.53975");
