// vybe-test: csharp/csharp_integer_arithmetic/subtraction_yields_negative_difference_when_subtrahend_is_larger
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((4 - 15).ToString(), "-11");
