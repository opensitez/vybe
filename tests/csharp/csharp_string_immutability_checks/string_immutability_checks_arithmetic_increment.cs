// vybe-test: csharp/csharp_string_immutability_checks/string_immutability_checks_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_string_immutability_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_immutability_checks
int seed = 18; __Check((seed + 1 > seed).ToString(), "True");
