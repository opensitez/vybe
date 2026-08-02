// vybe-test: csharp/csharp_checked_context_math/checked_context_math_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_checked_context_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// checked_context_math
int? maybe = 12; __Check((maybe.HasValue && maybe.Value == 12).ToString(), "True");
