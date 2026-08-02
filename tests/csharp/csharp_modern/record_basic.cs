// vybe-test: csharp/csharp_modern/record_basic
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y);
var p = new Point(3, 4);
__Check((p.X).ToString(), "3");
__Check((p.Y).ToString(), "4");
