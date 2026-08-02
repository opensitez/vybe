// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_unicode_basics
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
