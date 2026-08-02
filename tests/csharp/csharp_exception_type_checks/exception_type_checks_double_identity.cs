// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// exception_type_checks
double seed = 53; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
