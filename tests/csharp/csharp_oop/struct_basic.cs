// vybe-test: csharp/csharp_oop/struct_basic
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Point {
    public int X;
    public int Y;
    public Point(int x, int y) { X = x; Y = y; }
    public int Sum() { return X + Y; }
}
var p = new Point(3, 4);
__Check((p.Sum()).ToString(), "7");
