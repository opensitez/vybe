// vybe-test: csharp/csharp_pattern_constant_checks/pattern_constant_checks_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_pattern_constant_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_constant_checks
double seed = 40; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
