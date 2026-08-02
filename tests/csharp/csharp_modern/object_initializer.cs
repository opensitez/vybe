// vybe-test: csharp/csharp_modern/object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Point {
    public int X { get; set; }
    public int Y { get; set; }
}
var p = new Point { X = 10, Y = 20 };
__Check((p.X + p.Y).ToString(), "30");
