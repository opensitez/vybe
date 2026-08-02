// vybe-test: csharp/oop_advanced/constructor_chaining_this
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
    public Point() : this(0, 0) { }
    public Point(int x, int y) { X = x; Y = y; }
}
var a = new Point();
var b = new Point(5, 10);
__Check((a.X + "," + a.Y).ToString(), "0,0");
__Check((b.X + "," + b.Y).ToString(), "5,10");
