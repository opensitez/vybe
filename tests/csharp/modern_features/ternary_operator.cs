// vybe-test: csharp/modern_features/ternary_operator
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 10;
string result = x > 5 ? "big" : "small";
__Check((result).ToString(), "big");
