// vybe-test: csharp/csharp_pattern_is_checks/pattern_is_checks_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_pattern_is_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_is_checks
int seed = 41; int right = seed + 1; __Check((seed < right).ToString(), "True");
