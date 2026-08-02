// vybe-test: csharp/csharp_abstract_sealed/abstract_method_must_be_overridden_and_is_dispatched_polymorphically
// origin: languages/csharp/tests/csharp/test_csharp_abstract_sealed.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Shape { public abstract double Area(); }
class Circle : Shape {
    public double R;
    public override double Area() => System.Math.PI * R * R;
}
Shape s = new Circle { R = 0 };
__Check((s.Area()).ToString(), "0");
