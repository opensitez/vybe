// vybe-test: csharp/csharp_oop/abstract_class_and_override
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

using static __Harness;

var sq = new Square(5);
__P((sq.Area()).ToString());
__P((sq.Describe()).ToString());
__Check("25\nArea=25");

abstract class Shape {
    public abstract double Area();
    public string Describe() { return "Area=" + Area(); }
}

class Square : Shape {
    public double Side;
    public Square(double s) { Side = s; }
    public override double Area() { return Side * Side; }
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
