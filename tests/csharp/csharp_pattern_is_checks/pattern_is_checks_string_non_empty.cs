// vybe-test: csharp/csharp_pattern_is_checks/pattern_is_checks_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_pattern_is_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_is_checks
string feature = "pattern_is_checks"; __Check((feature.Length > 0).ToString(), "True");
