// vybe-test: csharp/csharp_pattern_constant_checks/pattern_constant_checks_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_pattern_constant_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_constant_checks
int? maybe = 40; __Check((maybe.HasValue && maybe.Value == 40).ToString(), "True");
