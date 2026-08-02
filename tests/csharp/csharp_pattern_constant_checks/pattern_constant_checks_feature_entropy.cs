// vybe-test: csharp/csharp_pattern_constant_checks/pattern_constant_checks_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_pattern_constant_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_constant_checks
string feature = "pattern_constant_checks:40"; __Check((feature.Length >= 1).ToString(), "True");
