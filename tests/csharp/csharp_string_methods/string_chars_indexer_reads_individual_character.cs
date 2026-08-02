// vybe-test: csharp/csharp_string_methods/string_chars_indexer_reads_individual_character
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("hello"[1]).ToString(), "e");
