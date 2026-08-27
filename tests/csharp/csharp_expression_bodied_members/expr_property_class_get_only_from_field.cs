// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_get_only_from_field
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P((System.Math.Round(new Circle().Area, 2)).ToString());
__Check("12.57");

class Circle { public double R = 2.0; public double Area => System.Math.PI * R * R; }

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
