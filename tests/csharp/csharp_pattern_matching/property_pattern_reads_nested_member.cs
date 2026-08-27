// vybe-test: csharp/csharp_pattern_matching/property_pattern_reads_nested_member
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

using static __Harness;

object r = new Rect { W=10, H=5 }
;
string size = r switch { Rect { W: > 8 } => "wide", _ => "narrow" }
;
__P((size).ToString());
__Check("wide");

class Rect { public int W, H; }

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
