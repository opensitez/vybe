// vybe-test: csharp/modern_features/tuple_basic
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var t = (1, "hello", true);
__Check((t.Item1).ToString(), "1");
__Check((t.Item2).ToString(), "hello");
__Check((t.Item3).ToString(), "True");
