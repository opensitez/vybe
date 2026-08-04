// vybe-test: csharp/csharp_patterns/factory_method
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

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

class Shape {
    public string Type;
    private Shape(string t) { Type = t; }
    public static Shape Circle() { return new Shape("circle"); }
    public static Shape Square() { return new Shape("square"); }
}
var c = Shape.Circle();
var s = Shape.Square();
__P((c.Type).ToString());
__P((s.Type).ToString());
__Check("circle\nsquare");
