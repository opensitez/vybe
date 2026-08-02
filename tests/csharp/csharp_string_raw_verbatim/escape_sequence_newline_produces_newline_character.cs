// vybe-test: csharp/csharp_string_raw_verbatim/escape_sequence_newline_produces_newline_character
// origin: languages/csharp/tests/csharp/test_csharp_string_raw_verbatim.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s="a\nb"; __Check((s.Length).ToString(), "3");
