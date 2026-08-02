// vybe-test: csharp/csharp_pattern_constant_checks/pattern_constant_checks_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_pattern_constant_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_constant_checks
string feature = "pattern_constant_checks"; __Check((feature[0] == feature[0]).ToString(), "True");
