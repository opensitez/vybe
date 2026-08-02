// vybe-test: csharp/csharp_string_immutability_checks/string_immutability_checks_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_string_immutability_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_immutability_checks
int seed = 18; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
