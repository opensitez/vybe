// vybe-test: csharp/csharp_string_raw_verbatim/unicode_escape_produces_correct_character
// origin: languages/csharp/tests/csharp/test_csharp_string_raw_verbatim.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char c='\u0041'; __Check((c).ToString(), "A");
