// vybe-test: csharp/csharp_integer_arithmetic/double_unary_minus_restores_positive_value
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 15; __Check((-(-value)).ToString(), "15");
