// vybe-test: csharp/csharp_equality_contracts/struct_equality_compares_field_values_when_equals_overridden
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

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

struct Point {
    public int X;
    public int Y;
    public bool Equals(Point other) { return X == other.X && Y == other.Y; }
}
var left = new Point { X = 2, Y = 3 };
var right = new Point { X = 2, Y = 3 };
__P((left.Equals(right)).ToString());
__Check("True");
