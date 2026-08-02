// vybe-test: csharp/csharp_oop/class_readonly_field
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Circle {
    public readonly double PI = 3.14159;
    public double Radius;
    public Circle(double r) { Radius = r; }
    public double Area() { return PI * Radius * Radius; }
}
var c = new Circle(10);
__Check((c.Area()).ToString(), "314.159");
