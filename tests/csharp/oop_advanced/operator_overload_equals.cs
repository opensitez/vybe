// vybe-test: csharp/oop_advanced/operator_overload_equals
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

var a = new Point(1, 2);
var b = new Point(1, 2);
var c = new Point(3, 4);
__P((a == b).ToString());
__P((a != c).ToString());
__Check("True\nTrue");

class Point {
    public int X { get; set; }
    public int Y { get; set; }
    public Point(int x, int y) { X = x; Y = y; }
    public static bool operator ==(Point a, Point b) {
        return a.X == b.X && a.Y == b.Y;
    }
    public static bool operator !=(Point a, Point b) {
        return !(a == b);
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
