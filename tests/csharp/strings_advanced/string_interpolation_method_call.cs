// vybe-test: csharp/strings_advanced/string_interpolation_method_call
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = "hello";
__Check(($"upper: {s.ToUpper()}").ToString(), "upper: HELLO");
