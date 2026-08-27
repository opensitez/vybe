// vybe-test: csharp/csharp_anonymous_type_inference/anonymous_type_infers_property_names_from_simple_identifier_expressions
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_type_inference.rs

using static __Harness;

int width = 4;
string label = "box";
var shape = new { width, label }
;
__P((shape.width).ToString());
__P((shape.label).ToString());
__Check("4\nbox");

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
