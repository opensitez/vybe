// vybe-test: csharp/modern_features/is_constant_pattern
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object obj = null;
__Check((obj is null).ToString(), "True");
obj = 42;
__Check((obj is 42).ToString(), "True");
__Check((obj is 43).ToString(), "False");
