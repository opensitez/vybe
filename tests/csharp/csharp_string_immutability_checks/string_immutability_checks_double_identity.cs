// vybe-test: csharp/csharp_string_immutability_checks/string_immutability_checks_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_string_immutability_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_immutability_checks
double seed = 18; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
