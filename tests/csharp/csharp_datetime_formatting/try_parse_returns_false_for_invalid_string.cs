// vybe-test: csharp/csharp_datetime_formatting/try_parse_returns_false_for_invalid_string
// origin: languages/csharp/tests/csharp/test_csharp_datetime_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.DateTime.TryParse("not-a-date", out _)).ToString(), "False");
