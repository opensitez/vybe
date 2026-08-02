// vybe-test: csharp/csharp_break_continue_surface/break_continue_surface_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_break_continue_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// break_continue_surface
int? maybe = 49; __Check((maybe.HasValue && maybe.Value == 49).ToString(), "True");
