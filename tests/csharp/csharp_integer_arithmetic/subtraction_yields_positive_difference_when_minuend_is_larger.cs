// vybe-test: csharp/csharp_integer_arithmetic/subtraction_yields_positive_difference_when_minuend_is_larger
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((15 - 4).ToString(), "11");
