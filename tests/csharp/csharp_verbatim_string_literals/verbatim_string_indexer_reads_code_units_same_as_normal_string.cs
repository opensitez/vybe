// vybe-test: csharp/csharp_verbatim_string_literals/verbatim_string_indexer_reads_code_units_same_as_normal_string
// origin: languages/csharp/tests/csharp/test_csharp_verbatim_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((@"abc"[1]).ToString(), "b");
