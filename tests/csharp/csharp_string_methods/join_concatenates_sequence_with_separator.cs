// vybe-test: csharp/csharp_string_methods/join_concatenates_sequence_with_separator
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Join("-", new[]{"a","b","c"})).ToString(), "a-b-c");
