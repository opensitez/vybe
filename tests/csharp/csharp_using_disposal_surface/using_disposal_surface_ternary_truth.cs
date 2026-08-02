// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// using_disposal_surface
int seed = 52; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
