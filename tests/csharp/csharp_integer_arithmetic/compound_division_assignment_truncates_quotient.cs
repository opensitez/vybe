// vybe-test: csharp/csharp_integer_arithmetic/compound_division_assignment_truncates_quotient
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 17; value /= 5; __Check((value).ToString(), "3");
