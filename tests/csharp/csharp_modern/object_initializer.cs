// vybe-test: csharp/csharp_modern/object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

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
}
var p = new Point { X = 10, Y = 20 };
__P((p.X + p.Y).ToString());
__Check("30");
