// vybe-test: csharp/csharp_nested_partial_types/nested_struct_can_be_created_from_outer_scope
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

using static __Harness;

var point = new Geometry.Point { X = 3, Y = 4 }
;
__P((point.X + point.Y).ToString());
__Check("7");

class Geometry {
    public struct Point {
        public int X;
        public int Y;
    }
}

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
