// vybe-test: csharp/csharp_interface_explicit_impl/explicit_impl_is_not_accessible_through_class_reference
// origin: languages/csharp/tests/csharp/test_csharp_interface_explicit_impl.rs

using static __Harness;

IArea shape = new Square { Side = 3 }
;
__P((shape.Area()).ToString());
__Check("9");

interface IArea { double Area(); }

class Square : IArea {
    public double Side;
    double IArea.Area() => Side * Side;
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
