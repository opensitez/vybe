// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_unicode_basics
string feature = "string_unicode_basics"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
