// vybe-test: csharp/csharp_verbatim_string_literals/verbatim_string_spans_multiple_lines_when_source_contains_newlines
// origin: languages/csharp/tests/csharp/test_csharp_verbatim_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((@"line1
line2").ToString(), "line1\nline2");
