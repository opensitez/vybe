// vybe-test: csharp/csharp_pattern_property/is_property_pattern_enum_field_match
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object o=new Paint{Hue=Color.Green}
;
__P((o is Paint{Hue:Color.Green}).ToString());
__Check("True");

enum Color { Red, Green }

class Paint { public Color Hue; }

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
