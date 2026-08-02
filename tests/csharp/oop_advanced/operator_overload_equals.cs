// vybe-test: csharp/oop_advanced/operator_overload_equals
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

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
var a = new Point(1, 2);
var b = new Point(1, 2);
var c = new Point(3, 4);
__Check((a == b).ToString(), "True");
__Check((a != c).ToString(), "True");
