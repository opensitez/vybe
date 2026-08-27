// vybe-test: csharp/csharp_structs_value_semantics/struct_method_can_compute_from_fields
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

var point = new Point { X = 4, Y = 6 }
;
__P((point.Sum()).ToString());
__Check("10");

struct Point { public int X; public int Y; public int Sum() { return X + Y; } }

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
