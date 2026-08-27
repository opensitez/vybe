// vybe-test: csharp/csharp_pattern_property/switch_expression_property_or_literal_kind_arms
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

string Label(object o)=>o switch{Msg{Kind:"err" or "fail"}=>"bad",_=>"ok"}
;
__P((Label(new Msg{Kind="fail"})).ToString());
__Check("bad");

class Msg { public string Kind; }

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
