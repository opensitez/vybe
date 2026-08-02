// vybe-test: csharp/csharp_char_type_semantics/char_literal_single_quote_denotes_code_unit
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char letter = 'A'; __Check((letter).ToString(), "A");
