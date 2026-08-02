// vybe-test: csharp/csharp_tuples_ranges/int_parse
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = int.Parse("42");
__Check((n + 8).ToString(), "50");
