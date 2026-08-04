// vybe-test: csharp/csharp_abstract_sealed/abstract_method_must_be_overridden_and_is_dispatched_polymorphically
// origin: languages/csharp/tests/csharp/test_csharp_abstract_sealed.rs

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

abstract class Shape { public abstract double Area(); }
class Circle : Shape {
    public double R;
    public override double Area() => System.Math.PI * R * R;
}
Shape s = new Circle { R = 0 };
__P((s.Area()).ToString());
__Check("0");
