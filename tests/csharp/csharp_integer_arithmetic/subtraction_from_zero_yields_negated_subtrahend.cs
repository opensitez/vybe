// vybe-test: csharp/csharp_integer_arithmetic/subtraction_from_zero_yields_negated_subtrahend
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((0 - 8).ToString(), "-8");
