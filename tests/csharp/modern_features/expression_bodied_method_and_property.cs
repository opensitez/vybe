// vybe-test: csharp/modern_features/expression_bodied_method_and_property
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

var c = new Circle(5);
__P(Math.Round(c.Area, 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("78.5");

class Circle {
    private double r;
    public Circle(double r) => this.r = r;
    public double Area => Math.PI * r * r;
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
