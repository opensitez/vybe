// vybe-test: csharp/csharp_verbatim_string_literals/verbatim_empty_string_has_zero_length
// origin: languages/csharp/tests/csharp/test_csharp_verbatim_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((@"".Length).ToString(), "0");
