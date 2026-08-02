// vybe-test: csharp/csharp_string_parsing/guid_try_parse_recognises_standard_format
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Guid.TryParse("550e8400-e29b-41d4-a716-446655440000",out _)).ToString(), "True");
