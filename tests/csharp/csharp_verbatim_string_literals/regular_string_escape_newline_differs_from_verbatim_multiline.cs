// vybe-test: csharp/csharp_verbatim_string_literals/regular_string_escape_newline_differs_from_verbatim_multiline
// origin: languages/csharp/tests/csharp/test_csharp_verbatim_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("a\nb").ToString(), "a\nb");
