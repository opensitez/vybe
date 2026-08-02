// vybe-test: csharp/csharp_integer_arithmetic/compound_subtraction_exceeding_value_flips_sign
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 4; value -= 9; __Check((value).ToString(), "-5");
