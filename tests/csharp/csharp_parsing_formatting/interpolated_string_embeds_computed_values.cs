// vybe-test: csharp/csharp_parsing_formatting/interpolated_string_embeds_computed_values
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(($"sum={2 + 3}").ToString(), "sum=5");
