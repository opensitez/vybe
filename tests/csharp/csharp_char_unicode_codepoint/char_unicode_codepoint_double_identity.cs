// vybe-test: csharp/csharp_char_unicode_codepoint/char_unicode_codepoint_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_char_unicode_codepoint.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_unicode_codepoint
double seed = 22; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
