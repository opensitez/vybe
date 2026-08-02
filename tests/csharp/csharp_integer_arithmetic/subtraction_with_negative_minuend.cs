// vybe-test: csharp/csharp_integer_arithmetic/subtraction_with_negative_minuend
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((-9 - 4).ToString(), "-13");
