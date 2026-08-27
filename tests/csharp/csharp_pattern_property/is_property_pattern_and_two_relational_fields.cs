// vybe-test: csharp/csharp_pattern_property/is_property_pattern_and_two_relational_fields
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object o=new Range{Lo=5,Hi=15}
;
__P((o is Range{Lo:>0,Hi:<20}).ToString());
__Check("True");

class Range { public int Lo; public int Hi; }

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
