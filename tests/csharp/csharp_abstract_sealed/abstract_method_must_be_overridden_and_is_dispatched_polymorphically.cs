// vybe-test: csharp/csharp_abstract_sealed/abstract_method_must_be_overridden_and_is_dispatched_polymorphically
// origin: languages/csharp/tests/csharp/test_csharp_abstract_sealed.rs

using static __Harness;

Shape s = new Circle { R = 0 }
;
__P((s.Area()).ToString());
__Check("0");

abstract class Shape { public abstract double Area(); }

class Circle : Shape {
    public double R;
    public override double Area() => System.Math.PI * R * R;
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
