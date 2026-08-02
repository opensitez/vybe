// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// exception_type_checks
int? maybe = null; int fallback = maybe ?? 53; __Check((fallback == 53).ToString(), "True");
