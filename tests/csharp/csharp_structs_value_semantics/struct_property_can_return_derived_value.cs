// vybe-test: csharp/csharp_structs_value_semantics/struct_property_can_return_derived_value
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

var rect = new Rect { W = 3, H = 5 }
;
__P((rect.Area).ToString());
__Check("15");

struct Rect { public int W { get; set; } public int H { get; set; } public int Area => W * H; }

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
