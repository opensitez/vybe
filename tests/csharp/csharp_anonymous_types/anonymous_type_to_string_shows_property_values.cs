// vybe-test: csharp/csharp_anonymous_types/anonymous_type_to_string_shows_property_values
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

using static __Harness;

var a=new{X=3,Y=4}
;
__P((a.ToString().Contains("X = 3")).ToString());
__Check("True");

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
