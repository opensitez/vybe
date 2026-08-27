// vybe-test: csharp/csharp_oop/struct_basic
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

using static __Harness;

var p = new Point(3, 4);
__P((p.Sum()).ToString());
__Check("7");

struct Point {
    public int X;
    public int Y;
    public Point(int x, int y) { X = x; Y = y; }
    public int Sum() { return X + Y; }
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
