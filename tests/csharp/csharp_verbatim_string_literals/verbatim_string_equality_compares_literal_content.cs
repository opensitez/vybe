// vybe-test: csharp/csharp_verbatim_string_literals/verbatim_string_equality_compares_literal_content
// origin: languages/csharp/tests/csharp/test_csharp_verbatim_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((@"x" == "x").ToString(), "True");
