// vybe-test: csharp/common_patterns/readonly_field
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Circle {
    public readonly double Pi = 3.14159;
    public double Radius;
    public Circle(double r) { Radius = r; }
    public double Area() { return Pi * Radius * Radius; }
}
var c = new Circle(1);
__Check((c.Pi).ToString(), "3.14159");
