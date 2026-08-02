// vybe-test: csharp/modern_features/record_equality
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Color(int R, int G, int B);
var c1 = new Color(255, 0, 0);
var c2 = new Color(255, 0, 0);
var c3 = new Color(0, 255, 0);
__Check((c1 == c2).ToString(), "True");
__Check((c1 == c3).ToString(), "False");
