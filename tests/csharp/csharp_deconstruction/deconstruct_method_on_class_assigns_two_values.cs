// vybe-test: csharp/csharp_deconstruction/deconstruct_method_on_class_assigns_two_values
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

using static __Harness;

var point = new Point(8, 13);
var (x, y) = point;
__P((x).ToString());
__P((y).ToString());
__Check("8\n13");

class Point {
    int x;
    int y;
    public Point(int x, int y) { this.x = x; this.y = y; }
    public void Deconstruct(out int xValue, out int yValue) {
        xValue = x;
        yValue = y;
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
