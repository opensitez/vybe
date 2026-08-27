// vybe-test: csharp/csharp_pattern_property/switch_expression_property_when_guard_on_fields
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

string Sign(object o)=>o switch{Pair{A:var x,B:var y} when x==y=>"eq",Pair{A:var x,B:var y}=>"neq",_=>"?"}
;
__P((Sign(new Pair{A=3,B=3})).ToString());
__Check("eq");

class Pair { public int A; public int B; }

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
