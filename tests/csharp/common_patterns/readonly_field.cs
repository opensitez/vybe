// vybe-test: csharp/common_patterns/readonly_field
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

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
    public readonly double Pi = 3.14159;
    public double Radius;
    public Circle(double r) { Radius = r; }
    public double Area() { return Pi * Radius * Radius; }
}
var c = new Circle(1);
__P((c.Pi).ToString());
__Check("3.14159");
