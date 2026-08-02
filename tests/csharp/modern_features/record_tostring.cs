// vybe-test: csharp/modern_features/record_tostring
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y);
var p = new Point(3, 4);
__Check((p).ToString(), "Point { X = 3, Y = 4 }");
