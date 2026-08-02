// vybe-test: csharp/csharp_verbatim_string_literals/verbatim_string_doubled_quote_embeds_single_quote
// origin: languages/csharp/tests/csharp/test_csharp_verbatim_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((@"say ""hi""").ToString(), "say \"hi\"");
