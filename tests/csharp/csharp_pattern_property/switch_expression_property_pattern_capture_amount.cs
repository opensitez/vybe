// vybe-test: csharp/csharp_pattern_property/switch_expression_property_pattern_capture_amount
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

int Read(object o)=>o switch{Wallet{Balance:var b}=>b,_=>-1}
;
__P((Read(new Wallet{Balance=42})).ToString());
__Check("42");

class Wallet { public int Balance; }

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
