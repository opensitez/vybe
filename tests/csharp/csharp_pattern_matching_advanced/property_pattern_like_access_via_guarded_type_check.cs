// vybe-test: csharp/csharp_pattern_matching_advanced/property_pattern_like_access_via_guarded_type_check
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

using static __Harness;

object item = new Point { X = 5, Y = 8 }
;
if (item is Point point && point.X == 5) __P((point.Y).ToString());
__Check("8");

class Point { public int X { get; set; } public int Y { get; set; } }

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
