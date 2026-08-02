// vybe-test: csharp/csharp_integer_arithmetic/reassignment_after_arithmetic_expression
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 1; value = value + 2 + 3; __Check((value).ToString(), "6");
