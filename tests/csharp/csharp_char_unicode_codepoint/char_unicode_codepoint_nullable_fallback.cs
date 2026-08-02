// vybe-test: csharp/csharp_char_unicode_codepoint/char_unicode_codepoint_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_char_unicode_codepoint.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_unicode_codepoint
int? maybe = null; int fallback = maybe ?? 22; __Check((fallback == 22).ToString(), "True");
