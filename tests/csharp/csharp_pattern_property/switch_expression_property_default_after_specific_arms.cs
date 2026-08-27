// vybe-test: csharp/csharp_pattern_property/switch_expression_property_default_after_specific_arms
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

string Name(object o)=>o switch{Token{Kind:"add"}=>"plus",Token{Kind:"sub"}=>"minus",_=>"other"}
;
__P((Name(new Token{Kind="mul"})).ToString());
__Check("other");

class Token { public string Kind; }

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
