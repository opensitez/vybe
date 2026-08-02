// vybe-test: csharp/csharp_comparison_operators_surface/comparison_operators_surface_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_comparison_operators_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// comparison_operators_surface
double seed = 13; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
