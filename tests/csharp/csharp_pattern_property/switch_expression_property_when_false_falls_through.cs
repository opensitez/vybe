// vybe-test: csharp/csharp_pattern_property/switch_expression_property_when_false_falls_through
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

string Flag(object o)=>o switch{Item{Q:var q} when q>10=>"big",Item{Q:var q}=>"small",_=>"?"}
;
__P((Flag(new Item{Q=3})).ToString());
__Check("small");

class Item { public int Q; }

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
