// vybe-test: csharp/oop_advanced/abstract_class_basic
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

var c = new Circle(5);
__P((c.Area()).ToString());
__P((c.Describe()).ToString());
__Check("78.5\nI am a shape");

abstract class Shape {
    public abstract double Area();
    public string Describe() { return "I am a shape"; }
}

class Circle : Shape {
    double radius;
    public Circle(double r) { radius = r; }
    public override double Area() { return 3.14 * radius * radius; }
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
