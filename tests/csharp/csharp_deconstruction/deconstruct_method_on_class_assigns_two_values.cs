// vybe-test: csharp/csharp_deconstruction/deconstruct_method_on_class_assigns_two_values
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

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
    int x;
    int y;
    public Point(int x, int y) { this.x = x; this.y = y; }
    public void Deconstruct(out int xValue, out int yValue) {
        xValue = x;
        yValue = y;
    }
}
var point = new Point(8, 13);
var (x, y) = point;
__P((x).ToString());
__P((y).ToString());
__Check("8\n13");
