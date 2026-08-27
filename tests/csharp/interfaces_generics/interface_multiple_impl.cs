// vybe-test: csharp/interfaces_generics/interface_multiple_impl
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

IShape c = new Circle { Radius = 10 }
;
IShape s = new Square { Side = 5 }
;
__P((c.Area()).ToString());
__P((s.Area()).ToString());
__Check("314\n25");

interface IShape {
    double Area();
}

class Circle : IShape {
    public double Radius;
    public double Area() { return 3.14 * Radius * Radius; }
}

class Square : IShape {
    public double Side;
    public double Area() { return Side * Side; }
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
