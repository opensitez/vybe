// vybe-test: csharp/modern_features/nested_ternary
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 50;
string cat = x < 0 ? "negative" : x == 0 ? "zero" : "positive";
__Check((cat).ToString(), "positive");
