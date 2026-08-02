// vybe-test: csharp/csharp_nullable_reference_checks/nullable_reference_checks_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_nullable_reference_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_reference_checks
int seed = 58; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
