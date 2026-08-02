// vybe-test: csharp/csharp_break_continue_surface/break_continue_surface_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_break_continue_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// break_continue_surface
double seed = 49; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
