// vybe-test: csharp/csharp_string_parsing/int_parse_converts_decimal_string
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((int.Parse("42")).ToString(), "42");
