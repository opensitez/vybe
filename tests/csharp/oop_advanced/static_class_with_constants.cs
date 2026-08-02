// vybe-test: csharp/oop_advanced/static_class_with_constants
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class Constants {
    public const double Pi = 3.14159;
    public const int MaxSize = 100;
}
__Check((Constants.Pi).ToString(), "3.14159");
__Check((Constants.MaxSize).ToString(), "100");
