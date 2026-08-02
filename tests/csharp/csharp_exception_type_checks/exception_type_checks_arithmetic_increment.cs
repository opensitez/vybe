// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// exception_type_checks
int seed = 53; __Check((seed + 1 > seed).ToString(), "True");
