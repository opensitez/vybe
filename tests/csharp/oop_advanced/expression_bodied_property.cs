// vybe-test: csharp/oop_advanced/expression_bodied_property
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Circle {
    public double Radius { get; set; }
    public double Area => 3.14 * Radius * Radius;
    public Circle(double r) { Radius = r; }
}
var c = new Circle(5);
__Check((c.Area).ToString(), "78.5");
