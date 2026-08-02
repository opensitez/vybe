// vybe-test: csharp/csharp_nullable_reference_checks/nullable_reference_checks_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_nullable_reference_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_reference_checks
int seed = 58; __Check((seed - seed == 0).ToString(), "True");
