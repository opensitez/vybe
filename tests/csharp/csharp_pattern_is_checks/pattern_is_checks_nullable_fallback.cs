// vybe-test: csharp/csharp_pattern_is_checks/pattern_is_checks_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_pattern_is_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_is_checks
int? maybe = null; int fallback = maybe ?? 41; __Check((fallback == 41).ToString(), "True");
