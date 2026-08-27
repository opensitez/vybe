// vybe-test: csharp/csharp_pattern_deconstruct/property_pattern_extracts_nested_property_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

using static __Harness;

object o = new Order { Amount = 100, IsPaid = true }
;
var label = o switch {
    Order { IsPaid: true, Amount: > 50 } => "big paid",
    Order { IsPaid: true }               => "small paid",
    _                                    => "unpaid"
}
;
__P((label).ToString());
__Check("big paid");

class Order { public int Amount; public bool IsPaid; }

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
