// vybe-test: csharp/csharp_oop/class_readonly_field
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

using static __Harness;

var c = new Circle(10);
__P((c.Area()).ToString());
__Check("314.159");

class Circle {
    public readonly double PI = 3.14159;
    public double Radius;
    public Circle(double r) { Radius = r; }
    public double Area() { return PI * Radius * Radius; }
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
