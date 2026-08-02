// vybe-test: csharp/oop_advanced/abstract_class_basic
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Shape {
    public abstract double Area();
    public string Describe() { return "I am a shape"; }
}
class Circle : Shape {
    double radius;
    public Circle(double r) { radius = r; }
    public override double Area() { return 3.14 * radius * radius; }
}
var c = new Circle(5);
__Check((c.Area()).ToString(), "78.5");
__Check((c.Describe()).ToString(), "I am a shape");
