// vybe-test: csharp/csharp_covariant_return_override/property_override_can_narrow_getter_return_type_covariantly
// origin: languages/csharp/tests/csharp/test_csharp_covariant_return_override.rs

using static __Harness;

Shape original = new Circle { Radius = 5 }
;
var copy = original.CloneShape();
__P((copy is Circle).ToString());
__Check("True");

class Shape { public virtual Shape CloneShape() { return new Shape(); } }

class Circle : Shape { public int Radius; public override Circle CloneShape() { return new Circle { Radius = Radius }; } }

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
