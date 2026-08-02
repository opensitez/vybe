// vybe-test: csharp/csharp_integer_arithmetic/multiplication_in_assignment_updates_product
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int factor = 5; factor = factor * 3; __Check((factor).ToString(), "15");
