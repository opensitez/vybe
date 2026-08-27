// vybe-test: csharp/common_patterns/readonly_field
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

var c = new Circle(1);
__P((c.Pi).ToString());
__Check("3.14159");

class Circle {
    public readonly double Pi = 3.14159;
    public double Radius;
    public Circle(double r) { Radius = r; }
    public double Area() { return Pi * Radius * Radius; }
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
