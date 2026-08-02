// vybe-test: csharp/csharp_tuple_projection_checks/tuple_projection_checks_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_tuple_projection_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_projection_checks
string feature = "tuple_projection_checks:36"; __Check((feature.Length >= 1).ToString(), "True");
