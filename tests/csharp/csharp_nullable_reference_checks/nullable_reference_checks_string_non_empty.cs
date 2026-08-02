// vybe-test: csharp/csharp_nullable_reference_checks/nullable_reference_checks_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_nullable_reference_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_reference_checks
string feature = "nullable_reference_checks"; __Check((feature.Length > 0).ToString(), "True");
