// vybe-test: csharp/csharp_integer_arithmetic/subtraction_of_self_yields_zero
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 42; __Check((value - value).ToString(), "0");
