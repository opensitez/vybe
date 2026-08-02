// vybe-test: csharp/csharp_integer_arithmetic/division_by_one_returns_dividend
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((42 / 1).ToString(), "42");
