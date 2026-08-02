// vybe-test: csharp/csharp_number_bases/underscore_separator_does_not_change_value
// origin: languages/csharp/tests/csharp/test_csharp_number_bases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((1_000_000).ToString(), "1000000");
