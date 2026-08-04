// vybe-test: csharp/more_classes/class_calling_another_class
// origin: languages/csharp/tests/csharp/test_more_classes.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

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
        var p1 = new Point(0, 0);
        var p2 = new Point(3, 4);
        var line = new Line(p1, p2);
        __P((line.start.x).ToString());
__Check("0");
