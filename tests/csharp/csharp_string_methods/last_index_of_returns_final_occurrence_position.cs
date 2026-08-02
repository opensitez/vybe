// vybe-test: csharp/csharp_string_methods/last_index_of_returns_final_occurrence_position
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("abcabc".LastIndexOf('b')).ToString(), "4");
