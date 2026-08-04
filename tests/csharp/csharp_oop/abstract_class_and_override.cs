// vybe-test: csharp/csharp_oop/abstract_class_and_override
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

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
__P((sq.Area()).ToString());
__P((sq.Describe()).ToString());
__Check("25\nArea=25");
