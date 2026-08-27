// vybe-test: csharp/csharp_virtual_dispatch_semantics/virtual_property_getter_dispatches_to_derived_accessor
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

using static __Harness;

Shape shape = new Triangle();
__P((shape.Sides).ToString());
__Check("3");

class Shape {
    public virtual int Sides { get { return 0; } }
}

class Triangle : Shape {
    public override int Sides { get { return 3; } }
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
