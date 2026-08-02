// vybe-test: csharp/csharp_char_unicode_codepoint/char_unicode_codepoint_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_char_unicode_codepoint.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_unicode_codepoint
int seed = 22; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
