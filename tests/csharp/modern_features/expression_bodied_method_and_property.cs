// vybe-test: csharp/modern_features/expression_bodied_method_and_property
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Circle {
    public double Radius { get; }
    public Circle(double r) => Radius = r;
    public double Area => 3.14 * Radius * Radius;
    public double Circumference() => 2 * 3.14 * Radius;
}
var c = new Circle(5);
__Check((c.Area).ToString(), "78.5");
__Check((c.Circumference()).ToString(), "31.4");
