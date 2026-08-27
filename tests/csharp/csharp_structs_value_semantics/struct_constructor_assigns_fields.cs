// vybe-test: csharp/csharp_structs_value_semantics/struct_constructor_assigns_fields
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

var point = new Point(2, 3);
__P((point.X + point.Y).ToString());
__Check("5");

struct Point { public int X; public int Y; public Point(int x, int y) { X = x; Y = y; } }

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
