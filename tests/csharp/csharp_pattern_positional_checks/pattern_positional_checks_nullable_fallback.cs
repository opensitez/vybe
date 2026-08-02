// vybe-test: csharp/csharp_pattern_positional_checks/pattern_positional_checks_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_pattern_positional_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_positional_checks
int? maybe = null; int fallback = maybe ?? 115; __Check((fallback == 115).ToString(), "True");
