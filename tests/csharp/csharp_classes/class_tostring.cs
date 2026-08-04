// vybe-test: csharp/csharp_classes/class_tostring
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

class Point {
    public int X;
    public int Y;
    public Point(int x, int y) { X = x; Y = y; }
    public override string ToString() { return "(" + X + ", " + Y + ")"; }
}
var p = new Point(3, 4);
__P((p.ToString()).ToString());
__Check("(3, 4)");
