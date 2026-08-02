// vybe-test: csharp/more_classes/class_calling_another_class
// origin: languages/csharp/tests/csharp/test_more_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
        __Check((line.start.x).ToString(), "0");
