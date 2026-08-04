// vybe-test: csharp/oop_advanced/constructor_chaining_this
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

class Point {
    public int X { get; set; }
    public int Y { get; set; }
    public Point() : this(0, 0) { }
    public Point(int x, int y) { X = x; Y = y; }
}
var a = new Point();
var b = new Point(5, 10);
__P((a.X + "," + a.Y).ToString());
__P((b.X + "," + b.Y).ToString());
__Check("0,0\n5,10");
