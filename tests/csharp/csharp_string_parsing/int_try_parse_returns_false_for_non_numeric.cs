// vybe-test: csharp/csharp_string_parsing/int_try_parse_returns_false_for_non_numeric
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((int.TryParse("abc",out _)).ToString(), "False");
