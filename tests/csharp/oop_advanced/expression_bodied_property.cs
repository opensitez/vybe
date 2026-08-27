// vybe-test: csharp/oop_advanced/expression_bodied_property
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

var c = new Circle(5);
__P((c.Area).ToString());
__Check("78.5");

class Circle {
    public double Radius { get; set; }
    public double Area => 3.14 * Radius * Radius;
    public Circle(double r) { Radius = r; }
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
