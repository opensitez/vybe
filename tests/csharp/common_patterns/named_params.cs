// vybe-test: csharp/common_patterns/named_params
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Rect {
    public static int Area(int width, int height) { return width * height; }
}
__Check((Rect.Area(width: 5, height: 3)).ToString(), "15");
__Check((Rect.Area(height: 10, width: 2)).ToString(), "20");
