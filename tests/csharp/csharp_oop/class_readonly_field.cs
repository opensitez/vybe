// vybe-test: csharp/csharp_oop/class_readonly_field
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

class Circle {
    public readonly double PI = 3.14159;
    public double Radius;
    public Circle(double r) { Radius = r; }
    public double Area() { return PI * Radius * Radius; }
}
var c = new Circle(10);
__P((c.Area()).ToString());
__Check("314.159");
