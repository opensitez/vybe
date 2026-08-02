// vybe-test: csharp/modern_features/tuple_named
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var p = (Name: "Alice", Age: 30);
__Check((p.Name).ToString(), "Alice");
__Check((p.Age).ToString(), "30");
