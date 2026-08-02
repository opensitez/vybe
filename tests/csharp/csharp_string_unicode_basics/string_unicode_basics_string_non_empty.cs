// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_unicode_basics
string feature = "string_unicode_basics"; __Check((feature.Length > 0).ToString(), "True");
