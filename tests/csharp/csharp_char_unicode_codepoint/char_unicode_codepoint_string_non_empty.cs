// vybe-test: csharp/csharp_char_unicode_codepoint/char_unicode_codepoint_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_char_unicode_codepoint.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_unicode_codepoint
string feature = "char_unicode_codepoint"; __Check((feature.Length > 0).ToString(), "True");
