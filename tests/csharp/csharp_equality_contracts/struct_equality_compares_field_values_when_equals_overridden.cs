// vybe-test: csharp/csharp_equality_contracts/struct_equality_compares_field_values_when_equals_overridden
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

using static __Harness;

var left = new Point { X = 2, Y = 3 }
;
var right = new Point { X = 2, Y = 3 }
;
__P((left.Equals(right)).ToString());
__Check("True");

struct Point {
    public int X;
    public int Y;
    public bool Equals(Point other) { return X == other.X && Y == other.Y; }
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
