// vybe-test: csharp/csharp_pattern_is_checks/pattern_is_checks_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_pattern_is_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_is_checks
string feature = "pattern_is_checks"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
