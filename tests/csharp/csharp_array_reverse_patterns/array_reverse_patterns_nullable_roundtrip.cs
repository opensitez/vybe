// vybe-test: csharp/csharp_array_reverse_patterns/array_reverse_patterns_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_array_reverse_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_reverse_patterns
int? maybe = 27; __Check((maybe.HasValue && maybe.Value == 27).ToString(), "True");
