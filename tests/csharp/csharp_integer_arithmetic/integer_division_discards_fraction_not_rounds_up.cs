// vybe-test: csharp/csharp_integer_arithmetic/integer_division_discards_fraction_not_rounds_up
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((9 / 4).ToString(), "2");
