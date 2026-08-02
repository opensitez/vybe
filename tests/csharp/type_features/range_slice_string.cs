// vybe-test: csharp/type_features/range_slice_string
// origin: languages/csharp/tests/csharp/test_type_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = "Hello World";
        var sub = s[0..5];
        __Check((sub).ToString(), "Hello");
