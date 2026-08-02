// vybe-test: csharp/csharp_nullable_reference_checks/nullable_reference_checks_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_nullable_reference_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_reference_checks
string feature = "nullable_reference_checks:58"; __Check((feature.Length >= 1).ToString(), "True");
