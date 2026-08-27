// vybe-test: csharp/csharp_classes/interface_implementation
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

using static __Harness;

var c = new Circle(5);
__P((c.Area()).ToString());
__Check("78.53975");

interface IShape {
    double Area();
}

class Circle : IShape {
    public double Radius;
    public Circle(double r) { Radius = r; }
    public double Area() { return 3.14159 * Radius * Radius; }
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
