// vybe-test: csharp/oop_advanced/expression_bodied_property
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

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
    public double Radius { get; set; }
    public double Area => 3.14 * Radius * Radius;
    public Circle(double r) { Radius = r; }
}
var c = new Circle(5);
__P((c.Area).ToString());
__Check("78.5");
