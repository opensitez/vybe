// vybe-test: csharp/csharp_pattern_property/is_property_pattern_string_var_capture
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object o=new Label{Text="go"}
;
if(o is Label{Text:var t}) __P((t.ToUpper()).ToString());
__Check("GO");

class Label { public string Text; }

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
