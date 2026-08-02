// vybe-test: csharp/common_patterns/const_field
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config {
    public const int MaxRetries = 3;
    public const string AppName = "MyApp";
}
__Check((Config.MaxRetries).ToString(), "3");
__Check((Config.AppName).ToString(), "MyApp");
