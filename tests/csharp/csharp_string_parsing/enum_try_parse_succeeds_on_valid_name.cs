// vybe-test: csharp/csharp_string_parsing/enum_try_parse_succeeds_on_valid_name
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Color{Red,Green,Blue}
__Check((System.Enum.TryParse<Color>("Green",out var c)).ToString(), "True");
__Check((c).ToString(), "Green");
