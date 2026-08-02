// vybe-test: csharp/csharp_break_continue_surface/break_continue_surface_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_break_continue_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// break_continue_surface
int seed = 49; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
