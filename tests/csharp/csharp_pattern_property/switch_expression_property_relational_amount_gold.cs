// vybe-test: csharp/csharp_pattern_property/switch_expression_property_relational_amount_gold
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

string Tier(object o)=>o switch{Bill{Amount:>=100}=>"gold",Bill{Amount:>=50}=>"silver",_=>"bronze"}
;
__P((Tier(new Bill{Amount=120})).ToString());
__Check("gold");

class Bill { public int Amount; }

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
