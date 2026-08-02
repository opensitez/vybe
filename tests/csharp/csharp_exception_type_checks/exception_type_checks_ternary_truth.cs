// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// exception_type_checks
int seed = 53; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
