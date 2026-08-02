// vybe-test: csharp/type_features/string_join_array
// origin: languages/csharp/tests/csharp/test_type_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var arr = new string[] {"a", "b", "c"};
        __Check((string.Join(",", arr)).ToString(), "a,b,c");
