// vybe-test: csharp/modern_features/expression_bodied_method_and_property
// origin: languages/csharp/tests/csharp/test_modern_features.rs

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

class Circle {
    public double Radius { get; }
    public Circle(double r) => Radius = r;
    public double Area => 3.14 * Radius * Radius;
    public double Circumference() => 2 * 3.14 * Radius;
}
var c = new Circle(5);
__P((c.Area).ToString());
__P((c.Circumference()).ToString());
__Check("78.5\n31.4");
