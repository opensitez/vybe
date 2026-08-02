// vybe-test: csharp/csharp_char_type_semantics/string_indexer_returns_char_at_position
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text = "cat"; __Check((text[1]).ToString(), "a");
