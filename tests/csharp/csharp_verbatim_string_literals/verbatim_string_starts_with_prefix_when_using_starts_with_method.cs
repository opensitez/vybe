// vybe-test: csharp/csharp_verbatim_string_literals/verbatim_string_starts_with_prefix_when_using_starts_with_method
// origin: languages/csharp/tests/csharp/test_csharp_verbatim_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((@"C:\data".StartsWith(@"C:")).ToString(), "True");
