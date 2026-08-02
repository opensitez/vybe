// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// jagged_array_patterns
int? maybe = null; int fallback = maybe ?? 28; __Check((fallback == 28).ToString(), "True");
