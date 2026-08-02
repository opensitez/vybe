// vybe-test: csharp/csharp_char_unicode_codepoint/char_unicode_codepoint_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_char_unicode_codepoint.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_unicode_codepoint
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
