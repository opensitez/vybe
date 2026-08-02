// vybe-test: csharp/csharp_integer_arithmetic/compound_multiplication_by_zero_zeros_variable
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 99; value *= 0; __Check((value).ToString(), "0");
