// vybe-test: csharp/csharp_parsing_formatting/bool_parse_reads_true_text
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((bool.Parse("True")).ToString(), "True");
