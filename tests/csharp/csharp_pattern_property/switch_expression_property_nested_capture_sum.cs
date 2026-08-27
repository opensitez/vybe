// vybe-test: csharp/csharp_pattern_property/switch_expression_property_nested_capture_sum
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

int Sum(object o)=>o switch{Wrap{Data:{A:var a,B:var b}}=>a+b,_=>0}
;
__P((Sum(new Wrap{Data=new Inner{A=6,B=7}})).ToString());
__Check("13");

class Inner { public int A; public int B; }

class Wrap { public Inner Data; }

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
