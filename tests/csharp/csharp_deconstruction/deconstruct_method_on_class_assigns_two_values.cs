// vybe-test: csharp/csharp_deconstruction/deconstruct_method_on_class_assigns_two_values
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Point {
    int x;
    int y;
    public Point(int x, int y) { this.x = x; this.y = y; }
    public void Deconstruct(out int xValue, out int yValue) {
        xValue = x;
        yValue = y;
    }
}
var point = new Point(8, 13);
var (x, y) = point;
__Check((x).ToString(), "8");
__Check((y).ToString(), "13");
