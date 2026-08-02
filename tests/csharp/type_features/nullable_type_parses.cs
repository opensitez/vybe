// vybe-test: csharp/type_features/nullable_type_parses
// origin: languages/csharp/tests/csharp/test_type_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string name = "hello";
        __Check((name).ToString(), "hello");
