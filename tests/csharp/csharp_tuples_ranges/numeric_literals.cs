// vybe-test: csharp/csharp_tuples_ranges/numeric_literals
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((0xFF).ToString(), "255");
__Check((0b1010).ToString(), "10");
__Check((1.5e2).ToString(), "150");
