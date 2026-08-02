// vybe-test: csharp/csharp_char_type_semantics/char_unicode_escape_specifies_code_point
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char ch = '\u0041'; __Check((ch).ToString(), "A");
