// vybe-test: csharp/type_features/multi_var_no_initializer
// origin: languages/csharp/tests/csharp/test_type_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string a = "hello", b = "world";
        __Check((a + " " + b).ToString(), "hello world");
