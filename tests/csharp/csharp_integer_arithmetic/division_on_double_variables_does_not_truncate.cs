// vybe-test: csharp/csharp_integer_arithmetic/division_on_double_variables_does_not_truncate
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double a = 9, b = 4; __Check((a / b).ToString(), "2.25");
