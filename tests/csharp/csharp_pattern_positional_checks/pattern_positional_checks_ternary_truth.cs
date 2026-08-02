// vybe-test: csharp/csharp_pattern_positional_checks/pattern_positional_checks_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_pattern_positional_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_positional_checks
int seed = 115; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
