// vybe-test: csharp/csharp_classes/interface_implementation
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

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

interface IShape {
    double Area();
}
class Circle : IShape {
    public double Radius;
    public Circle(double r) { Radius = r; }
    public double Area() { return 3.14159 * Radius * Radius; }
}
var c = new Circle(5);
__P((c.Area()).ToString());
__Check("78.53975");
