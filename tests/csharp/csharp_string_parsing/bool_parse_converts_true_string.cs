// vybe-test: csharp/csharp_string_parsing/bool_parse_converts_true_string
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((bool.Parse("True")).ToString(), "True");
