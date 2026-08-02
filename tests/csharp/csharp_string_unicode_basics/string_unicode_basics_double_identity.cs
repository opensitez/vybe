// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_unicode_basics
double seed = 19; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
