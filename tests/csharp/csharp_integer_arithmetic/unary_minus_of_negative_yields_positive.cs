// vybe-test: csharp/csharp_integer_arithmetic/unary_minus_of_negative_yields_positive
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((-(-8)).ToString(), "8");
