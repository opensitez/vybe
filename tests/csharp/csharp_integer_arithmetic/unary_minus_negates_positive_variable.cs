// vybe-test: csharp/csharp_integer_arithmetic/unary_minus_negates_positive_variable
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 42; __Check((-value).ToString(), "-42");
