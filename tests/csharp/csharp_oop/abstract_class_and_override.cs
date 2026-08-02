// vybe-test: csharp/csharp_oop/abstract_class_and_override
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Shape {
    public abstract double Area();
    public string Describe() { return "Area=" + Area(); }
}
class Square : Shape {
    public double Side;
    public Square(double s) { Side = s; }
    public override double Area() { return Side * Side; }
}
var sq = new Square(5);
__Check((sq.Area()).ToString(), "25");
__Check((sq.Describe()).ToString(), "Area=25");
