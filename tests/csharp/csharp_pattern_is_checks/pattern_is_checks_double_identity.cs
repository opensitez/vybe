// vybe-test: csharp/csharp_pattern_is_checks/pattern_is_checks_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_pattern_is_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_is_checks
double seed = 41; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
