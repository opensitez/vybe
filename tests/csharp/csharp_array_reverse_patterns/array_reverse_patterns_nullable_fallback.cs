// vybe-test: csharp/csharp_array_reverse_patterns/array_reverse_patterns_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_array_reverse_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_reverse_patterns
int? maybe = null; int fallback = maybe ?? 27; __Check((fallback == 27).ToString(), "True");
