// vybe-test: csharp/more_classes/class_calling_another_class
// origin: languages/csharp/tests/csharp/test_more_classes.rs

using static __Harness;

var p1 = new Point(0, 0);
var p2 = new Point(3, 4);
var line = new Line(p1, p2);
__P((line.start.x).ToString());
__Check("0");

class Point {
            public int x;
            public int y;
            public Point(int x, int y) { this.x = x; this.y = y; }
        }

class Line {
            public Point start;
            public Point endPt;
            public Line(Point s, Point e) { this.start = s; this.endPt = e; }
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
