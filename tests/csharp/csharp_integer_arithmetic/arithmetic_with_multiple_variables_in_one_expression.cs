// vybe-test: csharp/csharp_integer_arithmetic/arithmetic_with_multiple_variables_in_one_expression
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int a = 2; int b = 3; int c = 4; __Check((a + b * c).ToString(), "14");
