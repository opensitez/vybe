// vybe-test: csharp/modern_features/typeof_gettype
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((typeof(int).Name).ToString(), "Int32");
__Check((typeof(string).Name).ToString(), "String");
__Check((42.GetType().Name).ToString(), "Int32");
