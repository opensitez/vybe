// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// exception_type_checks
string feature = "exception_type_checks:53"; __Check((feature.Length >= 1).ToString(), "True");
