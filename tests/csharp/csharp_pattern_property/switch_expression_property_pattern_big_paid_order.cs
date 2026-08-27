// vybe-test: csharp/csharp_pattern_property/switch_expression_property_pattern_big_paid_order
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

string Label(object o)=>o switch{Order{Paid:true,Amount:>50}=>"big-paid",Order{Paid:true}=>"paid",_=>"open"}
;
__P((Label(new Order{Amount=100,Paid=true})).ToString());
__Check("big-paid");

class Order { public int Amount; public bool Paid; }

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
