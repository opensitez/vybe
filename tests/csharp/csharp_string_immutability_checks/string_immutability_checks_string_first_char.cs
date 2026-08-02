// vybe-test: csharp/csharp_string_immutability_checks/string_immutability_checks_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_string_immutability_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_immutability_checks
string feature = "string_immutability_checks"; __Check((feature[0] == feature[0]).ToString(), "True");
