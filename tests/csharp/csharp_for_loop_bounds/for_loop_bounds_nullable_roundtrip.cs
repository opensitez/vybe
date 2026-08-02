// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// for_loop_bounds
int? maybe = 45; __Check((maybe.HasValue && maybe.Value == 45).ToString(), "True");
