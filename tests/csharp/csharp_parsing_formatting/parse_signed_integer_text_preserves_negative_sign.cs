// vybe-test: csharp/csharp_parsing_formatting/parse_signed_integer_text_preserves_negative_sign
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((int.Parse("-9")).ToString(), "-9");
