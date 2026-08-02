// vybe-test: csharp/csharp_pattern_is_checks/pattern_is_checks_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_pattern_is_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_is_checks
string feature = "pattern_is_checks:41"; __Check((feature.Length >= 1).ToString(), "True");
