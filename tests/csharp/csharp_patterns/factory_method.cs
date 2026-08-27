// vybe-test: csharp/csharp_patterns/factory_method
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

using static __Harness;

var c = Shape.Circle();
var s = Shape.Square();
__P((c.Type).ToString());
__P((s.Type).ToString());
__Check("circle\nsquare");

class Shape {
    public string Type;
    private Shape(string t) { Type = t; }
    public static Shape Circle() { return new Shape("circle"); }
    public static Shape Square() { return new Shape("square"); }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
