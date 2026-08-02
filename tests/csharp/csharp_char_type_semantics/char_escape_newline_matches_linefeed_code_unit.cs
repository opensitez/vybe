// vybe-test: csharp/csharp_char_type_semantics/char_escape_newline_matches_linefeed_code_unit
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char ch = '\n'; __Check(((int)ch).ToString(), "10");
