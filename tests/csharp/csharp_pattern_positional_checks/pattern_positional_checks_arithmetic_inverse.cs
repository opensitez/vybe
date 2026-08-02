// vybe-test: csharp/csharp_pattern_positional_checks/pattern_positional_checks_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_pattern_positional_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_positional_checks
int seed = 115; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
