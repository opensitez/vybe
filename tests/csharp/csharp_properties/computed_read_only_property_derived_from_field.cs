// vybe-test: csharp/csharp_properties/computed_read_only_property_derived_from_field
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

using static __Harness;

__P((System.Math.Round(new Circle{Radius=0}.Area)).ToString());
__Check("0");

class Circle { public double Radius; public double Area => System.Math.PI * Radius * Radius; }

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
