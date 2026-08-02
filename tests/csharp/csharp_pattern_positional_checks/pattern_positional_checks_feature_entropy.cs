// vybe-test: csharp/csharp_pattern_positional_checks/pattern_positional_checks_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_pattern_positional_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_positional_checks
string feature = "pattern_positional_checks:115"; __Check((feature.Length >= 1).ToString(), "True");
