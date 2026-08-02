// vybe-test: csharp/csharp_integer_arithmetic/increment_on_negative_value_moves_toward_zero
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = -3; value++; __Check((value).ToString(), "-2");
