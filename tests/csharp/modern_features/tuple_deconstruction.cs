// vybe-test: csharp/modern_features/tuple_deconstruction
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (name, age) = ("Bob", 25);
__Check((name).ToString(), "Bob");
__Check((age).ToString(), "25");
