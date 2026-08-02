// vybe-test: csharp/csharp_string_comparison/ordinal_ignore_case_treats_same_letters_as_equal
// origin: languages/csharp/tests/csharp/test_csharp_string_comparison.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Compare("Hello","hello",System.StringComparison.OrdinalIgnoreCase) == 0).ToString(), "True");
