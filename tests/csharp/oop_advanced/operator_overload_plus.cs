// vybe-test: csharp/oop_advanced/operator_overload_plus
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

class Vector {
    public double X { get; set; }
    public double Y { get; set; }
    public Vector(double x, double y) { X = x; Y = y; }
    public static Vector operator +(Vector a, Vector b) {
        return new Vector(a.X + b.X, a.Y + b.Y);
    }
}
var a = new Vector(1, 2);
var b = new Vector(3, 4);
var c = a + b;
__P((c.X).ToString());
__P((c.Y).ToString());
__Check("4\n6");
