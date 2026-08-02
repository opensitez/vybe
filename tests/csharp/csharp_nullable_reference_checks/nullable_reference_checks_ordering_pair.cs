// vybe-test: csharp/csharp_nullable_reference_checks/nullable_reference_checks_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_nullable_reference_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_reference_checks
int seed = 58; int right = seed + 1; __Check((seed < right).ToString(), "True");
