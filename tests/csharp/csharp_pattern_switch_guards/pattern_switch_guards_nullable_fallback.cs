// vybe-test: csharp/csharp_pattern_switch_guards/pattern_switch_guards_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_guards.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_switch_guards
int? maybe = null; int fallback = maybe ?? 42; __Check((fallback == 42).ToString(), "True");
